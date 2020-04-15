import logging
import threading
from astropy.coordinates import SkyCoord, ICRS
from astropy import units as u
import numpy as np
import pythoncom
import win32com.client
from pyobs.mixins import FitsNamespaceMixin

from pyobs.interfaces import IFitsHeaderProvider, IMotion, IRaDecOffsets, ISyncTarget
from pyobs.modules import timeout
from pyobs.modules.telescope.basetelescope import BaseTelescope
from pyobs.utils.threads import LockWithAbort
from .com import com_device


log = logging.getLogger('pyobs')


class AscomTelescope(BaseTelescope, FitsNamespaceMixin, IFitsHeaderProvider, IRaDecOffsets, ISyncTarget):
    def __init__(self, device: str = None, *args, **kwargs):
        BaseTelescope.__init__(self, *args, **kwargs, motion_status_interfaces=['ITelescope'])

        # variables
        self._device = device

        # offsets in ra/dec
        self._offset_ra = 0
        self._offset_dec = 0

        # mixins
        FitsNamespaceMixin.__init__(self, *args, **kwargs)

    def open(self):
        """Open module.

        Raises:
            ValueError: If cannot connect to device.
        """
        BaseTelescope.open(self)

        # init COM
        pythoncom.CoInitialize()

        # do we need to chose a device?
        if not self._device:
            x = win32com.client.Dispatch("ASCOM.Utilities.Chooser")
            x.DeviceType = 'Telescope'
            self._device = x.Choose(None)
            log.info('Selected telescope "%s".', self._device)

        # open connection
        device = win32com.client.Dispatch(self._device)
        if device.Connected:
            log.info('Telescope was already connected.')
        else:
            device.Connected = True
            if device.Connected:
                log.info('Connected to telescope.')
            else:
                raise ValueError('Unable to connect to telescope.')

        # finish COM
        pythoncom.CoUninitialize()

    def close(self):
        """Clode module."""
        BaseTelescope.close(self)

        # get device
        with com_device(self._device) as device:
            # close connection
            if device.Connected:
                log.info('Disconnecting from telescope...')
                device.Connected = False

    @timeout(60000)
    def init(self, *args, **kwargs):
        """Initialize telescope.

        Raises:
            ValueError: If telescope could not be initialized.
        """

        # acquire lock
        with LockWithAbort(self._lock_moving, self._abort_move):
            # just move telescope to alt=15, az=180
            self._move_altaz(15, 180, self._abort_move)

    @timeout(60000)
    def park(self, *args, **kwargs):
        """Park telescope.

        Raises:
            ValueError: If telescope could not be parked.
        """

        # acquire lock
        with LockWithAbort(self._lock_moving, self._abort_move):
            # just move telescope to alt=15, az=180
            self._move_altaz(15, 180, self._abort_move, final_state=IMotion.Status.PARKED)

    @timeout(60000)
    def _move_altaz(self, alt: float, az: float, abort_event: threading.Event,
                    final_state: IMotion.Status = IMotion.Status.POSITIONED):
        """Actually moves to given coordinates. Must be implemented by derived classes.

        Args:
            alt: Alt in deg to move to.
            az: Az in deg to move to.
            abort_event: Event that gets triggered when movement should be aborted.
            final_state: Motion state to set after finished moving.

        Raises:
            Exception: On error.
        """

        # reset offsets
        self._offset_ra, self._offset_dec = 0, 0

        # get device
        with com_device(self._device) as device:
            # start slewing
            self._change_motion_status(IMotion.Status.SLEWING)
            log.info("Moving telescope to Alt=%.3f°, Az=%.3f°...", alt, az)
            device.Tracking = False
            device.SlewToAltAzAsync(az, alt)

            # wait for it
            while device.Slewing:
                abort_event.wait(1)

            # finish slewing
            device.Tracking = False
            self._change_motion_status(final_state)
            log.info('Reached destination')

    def _move_radec(self, ra: float, dec: float, abort_event: threading.Event):
        """Actually starts tracking on given coordinates. Must be implemented by derived classes.

        Args:
            ra: RA in deg to track.
            dec: Dec in deg to track.
            abort_event: Event that gets triggered when movement should be aborted.

        Raises:
            Exception: On any error.
        """

        # reset offsets
        self._offset_ra, self._offset_dec = 0, 0

        # to skycoords
        ra_dec = SkyCoord(ra * u.deg, dec * u.deg, frame=ICRS)

        # get device
        with com_device(self._device) as device:
            # start slewing
            self._change_motion_status(IMotion.Status.SLEWING)
            log.info("Moving telescope to RA=%s (%.5f°), Dec=%s (%.5f°)...",
                     ra_dec.ra.to_string(sep=':', unit=u.hour, pad=True), ra,
                     ra_dec.dec.to_string(sep=':', unit=u.deg, pad=True), dec)
            device.Tracking = True
            device.SlewToCoordinatesAsync(ra / 15., dec)

            # wait for it
            while device.Slewing:
                abort_event.wait(1)

            # finish slewing
            self._change_motion_status(IMotion.Status.TRACKING)

    @timeout(10000)
    def set_radec_offsets(self, dra: float, ddec: float, *args, **kwargs):
        """Move an RA/Dec offset.

        Args:
            dra: RA offset in degrees.
            ddec: Dec offset in degrees.

        Raises:
            ValueError: If offset could not be set.
        """

        # acquire lock
        with LockWithAbort(self._lock_moving, self._abort_move):
            # get device
            with com_device(self._device) as device:
                # start slewing
                self._change_motion_status(IMotion.Status.SLEWING)
                log.info('Setting telescope offsets to dRA=%.2f", dDec=%.2f"...', dra * 3600., ddec * 3600.)

                # get current coordinates (with old offsets)
                ra, dec = self.get_radec()

                # store offsets
                self._offset_ra = dra
                self._offset_dec = ddec

                # add offset
                ra += float(self._offset_ra / np.cos(np.radians(dec)))
                dec += float(self._offset_dec)

                # start slewing
                device.Tracking = True
                device.SlewToCoordinatesAsync(ra / 15., dec)

                # wait for it
                while device.Slewing:
                    self._abort_move.wait(1)

                # finish slewing
                self._change_motion_status(IMotion.Status.TRACKING)
                log.info('Reached destination.')

    def get_radec_offsets(self, *args, **kwargs) -> (float, float):
        """Get RA/Dec offset.

        Returns:
            Tuple with RA and Dec offsets.
        """
        return self._offset_ra, self._offset_dec

    def get_motion_status(self, interface: str = None, *args, **kwargs) -> IMotion.Status:
        """Returns current motion status.

        Args:
            interface: Name of interface to get status for, or None.

        Returns:
            A string from the Status enumerator.

        Raises:
            KeyError: If interface is not known.
        """

        # get device
        with com_device(self._device) as device:
            # what status are we in?
            if device.Tracking:
                return IMotion.Status.TRACKING.value
            elif device.Slewing:
                return IMotion.Status.SLEWING.value
            elif device.AtPark:
                return IMotion.Status.PARKED.value
            else:
                return IMotion.Status.IDLE.value

    def get_radec(self, *args, **kwargs) -> (float, float):
        """Returns current RA and Dec.

        Returns:
            Tuple of current RA and Dec in degrees.
        """

        # get device
        with com_device(self._device) as device:
            ra_off = self._offset_ra / np.cos(np.radians(device.Declination))
            return float(device.RightAscension * 15 - ra_off), float(device.Declination - self._offset_dec)

    def get_altaz(self, *args, **kwargs) -> (float, float):
        """Returns current Alt and Az.

        Returns:
            Tuple of current Alt and Az in degrees.
        """

        # get device
        with com_device(self._device) as device:
            # correct azimuth, Autoslew returns it as E of S (but slews to azimuth with E of N)... *sigh*
            az = device.Azimuth + 180
            if az > 360.:
                az -= 360

            # create sky coordinates
            return device.Altitude, az

    def stop_motion(self, device: str = None, *args, **kwargs):
        """Stop the motion.

        Args:
            device: Name of device to stop, or None for all.
        """

        # get device
        with com_device(self._device) as device:
            # stop telescope
            device.AbortSlew()
            device.Tracking = False

    def is_ready(self, *args, **kwargs) -> bool:
        """Returns the device is "ready", whatever that means for the specific device.

        Returns:
            Whether device is ready
        """
        return True

    def sync_target(self, *args, **kwargs):
        """Synchronize telescope on current target using current offsets."""

        # get current RA/Dec without offsets
        ra, dec = self.get_radec()

        # get device
        with com_device(self._device) as device:
            # sync
            device.SyncToCoordinates(ra / 15., dec)

    def get_fits_headers(self, namespaces: list = None, *args, **kwargs) -> dict:
        """Returns FITS header for the current status of this module.

        Args:
            namespaces: If given, only return FITS headers for the given namespaces.

        Returns:
            Dictionary containing FITS headers.
        """

        # get headers from base
        hdr = BaseTelescope.get_fits_headers(self)

        # get offsets
        ra_off, dec_off = self.get_radec_offsets()

        # define values to request
        hdr['RAOFF'] = (ra_off, 'RA offset [deg]')
        hdr['DECOFF'] = (dec_off, 'Dec offset [deg]')

        # return it
        return self._filter_fits_namespace(hdr, namespaces, **kwargs)


__all__ = ['AscomTelescope']

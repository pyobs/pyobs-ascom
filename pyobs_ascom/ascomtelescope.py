import logging
import threading
from astropy.coordinates import SkyCoord
from astropy import units as u
import numpy as np
import pythoncom
import win32com.client

from pyobs.interfaces import IFitsHeaderProvider, IMotion, IEquatorialMount
from pyobs.modules import timeout
from pyobs.modules.telescope.basetelescope import BaseTelescope
from pyobs.utils.threads import LockWithAbort
from pyobs.utils.time import Time
from .com import  com_device


log = logging.getLogger('pyobs')


class AscomTelescope(BaseTelescope, IFitsHeaderProvider, IEquatorialMount):
    def __init__(self, device: str = None, *args, **kwargs):
        BaseTelescope.__init__(self, *args, **kwargs, motion_status_interfaces=['ITelescope'])

        # variables
        self._device = device

        # offsets in ra/dec
        self._offset_ra = 0
        self._offset_dec = 0

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
        pythoncom.CoInitialize()

    def close(self):
        """Clode module."""
        BaseTelescope.close(self)

        # get device
        with com_device(self._device) as device:
            # close connection
            if device.Connected:
                log.info('Disconnecting from telescope...')
                device.Connected = False

    def init(self, *args, **kwargs):
        """Initialize telescope.

        Raises:
            ValueError: If telescope could not be initialized.
        """
        pass

    def park(self, *args, **kwargs):
        """Park telescope.

        Raises:
            ValueError: If telescope could not be parked.
        """
        pass

    def __move(self, ra: float, dec: float, tracking: bool, abort_event: threading.Event):
        """Move to given RA/Dec.

        Args:
            ra: RA in deg to track.
            dec: Dec in deg to track.
            tracking: Whether to start tracking.
            abort_event: Event that gets triggered when movement should be aborted.

        Raises:
            Exception: On any error.
        """

        # get current coordinates
        cur_ra, cur_dec = self.get_radec()

        # add offset
        ra += self._offset_ra * np.cos(cur_dec)
        dec += self._offset_dec

        # get device
        with com_device(self._device) as device:
            # start slewing
            self._change_motion_status(IMotion.Status.SLEWING)
            log.info("Moving telescope to RA=%.2f, Dec=%.2f...", ra, dec)
            device.SlewToCoordinates(ra / 15., dec)
            device.Tracking = tracking

            # wait for it
            while device.Slewing:
                abort_event.wait(1)

            # finish slewing
            self._change_motion_status(IMotion.Status.TRACKING if tracking else IMotion.Status.POSITIONED)
            log.info('Reached destination')

    def _track_radec(self, ra: float, dec: float, abort_event: threading.Event):
        """Actually starts tracking on given coordinates. Must be implemented by derived classes.

        Args:
            ra: RA in deg to track.
            dec: Dec in deg to track.
            abort_event: Event that gets triggered when movement should be aborted.

        Raises:
            Exception: On any error.
        """

        # move telescope
        self.__move(ra, dec, True, abort_event)

    @timeout(60000)
    def _move_altaz(self, alt: float, az: float, abort_event: threading.Event):
        """Actually moves to given coordinates. Must be implemented by derived classes.

        Args:
            alt: Alt in deg to move to.
            az: Az in deg to move to.
            abort_event: Event that gets triggered when movement should be aborted.

        Raises:
            Exception: On error.
        """

        # alt/az coordinates to ra/dec
        coords = SkyCoord(alt=alt * u.degree, az=az * u.degree, obstime=Time.now(),
                          location=self.location, frame='altaz')
        icrs = coords.icrs

        # move
        self.__move(icrs.ra.degree, icrs.dec.degree, False, abort_event)

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
            # start slewing
            self._change_motion_status(IMotion.Status.SLEWING)
            log.info("Setting telescope offsets to dRA=%.2f, dDec=%.2f...", dra, ddec)

            # get current coordinates
            ra, dec = self.get_radec()

            # set offsets
            self._offset_ra = dra
            self._offset_dec = ddec

            # move
            self.__move(ra, dec, True, self._abort_move)

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
            # alt/az coordinates to ra/dec
            coords = SkyCoord(alt=device.Altitude * u.degree, az=device.Azimuth * u.degree, obstime=Time.now(),
                              location=self.location, frame='altaz')
            icrs = coords.icrs

            # return RA/Dec
            return float(icrs.ra.degree - self._offset_ra * np.cos(self._offset_dec)), \
                   float(icrs.dec.degree - self._offset_dec)

    def get_altaz(self, *args, **kwargs) -> (float, float):
        """Returns current Alt and Az.

        Returns:
            Tuple of current Alt and Az in degrees.
        """

        # get device
        with com_device(self._device) as device:
            # create sky coordinates
            return device.Altitude, device.Azimuth

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


__all__ = ['AscomTelescope']

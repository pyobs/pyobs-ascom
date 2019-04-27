import logging
import threading
from astropy.coordinates import SkyCoord
from astropy import units as u
import pythoncom
import win32com.client

from pyobs.interfaces import IFitsHeaderProvider, IMotion
from pyobs.modules import timeout
from pyobs.modules.telescope.basetelescope import BaseTelescope
from pyobs.utils.time import Time
from .com import  com_device


log = logging.getLogger('pyobs')


class AscomTelescope(BaseTelescope, IFitsHeaderProvider):
    def __init__(self, device: str = None, *args, **kwargs):
        BaseTelescope.__init__(self, *args, **kwargs)

        # variables
        self._device = device

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

    @timeout(60000)
    def init(self, *args, **kwargs) -> bool:
        """Initialize telescope.

        Returns:
            (bool) Success
        """
        raise NotImplementedError

    @timeout(60000)
    def park(self, *args, **kwargs) -> bool:
        """Park telescope.

        Returns:
            (bool) Success
        """
        raise NotImplementedError

    def _track(self, ra: float, dec: float, abort_event: threading.Event):
        """Actually starts tracking on given coordinates. Must be implemented by derived classes.

        Args:
            ra: RA in deg to track.
            dec: Dec in deg to track.
            abort_event: Event that gets triggered when movement should be aborted.

        Raises:
            Exception: On any error.
        """

        """starts tracking on given coordinates"""
        
        # get device
        with com_device(self._device) as device:
            # start slewing
            self._change_motion_status(IMotion.Status.SLEWING)
            log.info("Moving telescope to RA=%.2f, Dec=%.2f...", ra, dec)
            device.Tracking = True
            device.SlewToCoordinates(ra / 15., dec)

            # finish slewing
            self._change_motion_status(IMotion.Status.TRACKING)
            log.info('Reached destination')

    @timeout(60000)
    def _move(self, alt: float, az: float, abort_event: threading.Event):
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

        # track
        self._track(icrs.ra.degree, icrs.dec.degree, abort_event)

    @timeout(10000)
    def offset(self, dalt: float, daz: float, *args, **kwargs) -> bool:
        """Move an Alt/Az offset, which will be reset on next call of track.

        Args:
            dalt: Altitude offset in degrees.
            daz: Azimuth offset in degrees.
        """
        raise NotImplementedError

    def reset_offset(self, *args, **kwargs) -> bool:
        """Reset Alt/Az offset."""
        raise NotImplementedError

    def get_motion_status(self, device: str = None) -> IMotion.Status:
        """Returns current motion status.

        Args:
            device: Name of device to get status for, or None.

        Returns:
            A string from the Status enumerator.
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

    def get_ra_dec(self) -> (float, float):
        """Returns current RA and Dec.

        Returns:
            Tuple of current RA and Dec in degrees.
        """

        # get device
        with com_device(self._device) as device:
            # create sky coordinates
            return device.RightAscension * 15, device.Declination

    def get_alt_az(self) -> (float, float):
        """Returns current Alt and Az.

        Returns:
            Tuple of current Alt and Az in degrees.
        """

        # get device
        with com_device(self._device) as device:
            # create sky coordinates
            return device.Altitude, device.Azimuth


__all__ = ['AscomTelescope']

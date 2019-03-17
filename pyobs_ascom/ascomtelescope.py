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

log = logging.getLogger('pyobs')


class AscomTelescope(BaseTelescope, IFitsHeaderProvider):
    def __init__(self, device: str, *args, **kwargs):
        BaseTelescope.__init__(self, *args, **kwargs)

        # variables
        self._device = device
        self._telescope = None

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

        # open connection
        self._telescope = win32com.client.Dispatch(self._device)
        if self._telescope.Connected:
            log.info('Telescope was already connected.')
        else:
            self._telescope.Connected = True
            if self._telescope.Connected:
                log.info('Connected to telescope.')
            else:
                raise ValueError('Unable to connect to telescope.')

    def close(self):
        """Clode module."""
        BaseTelescope.close(self)

        # close connection
        if self._telescope.Connected:
            log.info('Disconnecting from telescope...')
            self._telescope.Connected = False

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
        # init COM in thread
        pythoncom.CoInitialize()

        # start slewing
        log.info("Moving telescope to RA=%.2f, Dec=%.2f...", ra, dec)
        self._telescope.Tracking = True
        self._telescope.SlewToCoordinates(ra / 15., dec)

        # finish slewing
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
                          location=self.environment.location, frame='altaz')
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

    def get_motion_status(self) -> str:
        """Returns current motion status.

        Returns:
            A string from the Status enumerator.
        """

        # init COM in thread
        pythoncom.CoInitialize()

        # what status are we in?
        if self._telescope.Tracking:
            return IMotion.Status.TRACKING.value
        elif self._telescope.Slewing:
            return IMotion.Status.SLEWING.value
        elif self._telescope.AtPark:
            return IMotion.Status.PARKED.value
        else:
            return IMotion.Status.IDLE.value

    def get_ra_dec(self) -> (float, float):
        """Returns current RA and Dec.

        Returns:
            Tuple of current RA and Dec in degrees.
        """

        # init COM in thread
        pythoncom.CoInitialize()

        # create sky coordinates
        return self._telescope.RightAscension * 15, self._telescope.Declination

    def get_alt_az(self) -> (float, float):
        """Returns current Alt and Az.

        Returns:
            Tuple of current Alt and Az in degrees.
        """

        try:
            # init COM in thread
            pythoncom.CoInitialize()

            # create sky coordinates
            return self._telescope.Altitude, self._telescope.Azimuth
        except:
            log.exception("Error")


__all__ = ['AscomTelescope']

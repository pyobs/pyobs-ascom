import logging
from astropy.coordinates import SkyCoord
from astropy import units as u
import pythoncom
import win32com.client

from pytel.interfaces import IFitsHeaderProvider
from pytel.modules import timeout
from pytel.modules.telescope.basetelescope import BaseTelescope


class AscomTelescope(BaseTelescope, IFitsHeaderProvider):
    def __init__(self, *args, **kwargs):
        BaseTelescope.__init__(self, *args, **kwargs)

        # variables
        self._telescope = None

    def open(self):
        # init COM
        pythoncom.CoInitialize()

        # do we need to chose a device?
        device = self.config['device']
        if not device:
            x = win32com.client.Dispatch("ASCOM.Utilities.Chooser")
            x.DeviceType = 'Telescope'
            device = x.Choose(None)

        # open connection
        self._telescope = win32com.client.Dispatch(device)
        if self._telescope.Connected:
            logging.info('Telescope was already connected.')
        else:
            self._telescope.Connected = True
            if self._telescope.Connected:
                logging.info('Connected to telescope.')
            else:
                logging.info('Unable to connect to telescope.')
                raise ValueError('Could not connect to telescope.')

    def close(self):
        # close connection
        if self._telescope.Connected:
            logging.info('Disconnecting from telescope...')
            self._telescope.Connected = False

    @classmethod
    def default_config(cls):
        cfg = super(AscomTelescope, cls).default_config()
        cfg['device'] = None
        cfg['focuser'] = None
        return cfg

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

    @timeout(60000)
    def track(self, ra: float, dec: float, *args, **kwargs) -> bool:
        """starts tracking on given coordinates"""
        # init COM in thread
        pythoncom.CoInitialize()

        # start slewing
        logging.info("Moving telescope to RA=%.2f, Dec=%.2f...", ra, dec)
        self._telescope.Tracking = True
        self._telescope.SlewToCoordinates(ra / 15., dec)

        # finish slewing
        logging.info('Reached destination')
        return True

    @timeout(60000)
    def move(self, alt: float, az: float, *args, **kwargs) -> bool:
        """moves to given coordinates"""
        raise NotImplementedError

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

    def test(self):
        pass

    def status(self, *args, **kwargs) -> dict:
        """returns current status"""
        # init COM in thread
        pythoncom.CoInitialize()

        # get status
        s = super().status(*args, **kwargs)

        # telescope
        if self._telescope.Connected:
            status = 'tracking' if self._telescope.Tracking else 'slewing' if self._telescope.Slewing else 'idle'
            s['telescope']['status'] = status
            s['telescope']['ra'] = self._telescope.RightAscension * 15.0
            s['telescope']['dec'] = self._telescope.Declination
            s['telescope']['alt'] = self._telescope.Altitude
            s['telescope']['az'] = self._telescope.Azimuth

        # finished
        return s

    def get_fits_headers(self, *args, **kwargs) -> dict:
        """get FITS header for the saved status of the telescope"""
        # init COM in thread
        pythoncom.CoInitialize()

        # create sky coordinates
        c = SkyCoord(ra=self._telescope.RightAscension * u.hour, dec=self._telescope.Declination * u.deg, frame='icrs')

        # return header
        return {
            'CRVAL1': (c.ra.deg, 'Right ascension of telescope [degrees]'),
            'CRVAL2': (c.dec.deg, 'Declination of telescope [degrees]'),
            'TEL-RA': (c.ra.deg, 'Right ascension of telescope [degrees]'),
            'TEL-DEC': (c.dec.deg, 'Declination of telescope [degrees]'),
            'TEL-ZD': (90. - self._telescope.Altitude, 'Telescope zenith distance [degrees]'),
            'TEL-ALT': (self._telescope.Altitude, 'Telescope altitude [degrees]'),
            'TEL-AZ': (self._telescope.Azimuth, 'Telescope azimuth [degrees]'),
            'RA': (c.ra.to_string(sep=':', unit=u.hour, pad=True), 'Right ascension of telescope'),
            'DEC': (c.dec.to_string(sep=':', unit=u.deg, pad=True), 'Declination of telescope')
        }


__all__ = ['AscomTelescope']

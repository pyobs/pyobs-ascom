import logging
import time
import pythoncom
import win32com.client

from pyobs import PyObsModule
from pyobs.interfaces import IFocuser, IFitsHeaderProvider, IMotion
from pyobs.modules import timeout


log = logging.getLogger(__name__)


class AscomFocuser(PyObsModule, IFocuser, IFitsHeaderProvider):
    def __init__(self, device: str, *args, **kwargs):
        PyObsModule.__init__(self, *args, **kwargs)

        # variables
        self._device = device
        self._focuser = None

    def open(self):
        """Open module."""
        PyObsModule.open(self)

        # init COM
        pythoncom.CoInitialize()

        # init focuser
        self._focuser = win32com.client.Dispatch(self._device)
        if self._focuser.Connected:
            log.info('Focuser was already connected.')
        else:
            self._focuser.Connected = True
            if self._focuser.Connected:
                log.info('Connected to focuser.')
            else:
                raise ValueError('Unable to connect to focuser.')

    def close(self):
        """Close module."""
        PyObsModule.close(self)

        # close connection
        if self._focuser.Connected:
            log.info('Disconnecting from focuser...')
            self._focuser.Connected = False

    def get_fits_headers(self, *args, **kwargs) -> dict:
        """Returns FITS header for the current status of the telescope.

        Returns:
            Dictionary containing FITS headers.
        """
        # init COM in thread
        pythoncom.CoInitialize()

        # return header
        return {
            'TEL-FOCU': (self._focuser.Position / self._focuser.StepSize, 'Focus of telescope [mm]')
        }

    @timeout(60000)
    def set_focus(self, focus: float, *args, **kwargs):
        """Sets new focus.

        Args:
            focus: New focus value.
        """

        # init COM in thread
        pythoncom.CoInitialize()

        # calculating new focus and move it
        log.info('Moving focus to %.2fmm...', focus)
        foc = int(focus * self._focuser.StepSize)
        self._focuser.Move(foc)

        # wait for it
        while self._focuser.IsMoving:
            time.sleep(0.1)

        # finished
        log.info('Reached new focus of %.2mm.', self._focuser.Position / self._focuser.StepSize)

    def get_focus(self, *args, **kwargs) -> float:
        """Return current focus.

        Returns:
            Current focus.
        """

        # init COM in thread
        pythoncom.CoInitialize()

        # return current focus
        return self._focuser.Position / self._focuser.StepSize

    def get_motion_status(self, device: str = None) -> IMotion.Status:
        pass


__all__ = ['AscomFocuser']

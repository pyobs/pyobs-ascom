import logging
import time
import pythoncom
import win32com.client

from pyobs import PyObsModule
from pyobs.interfaces import IFocuser, IFitsHeaderProvider, IMotion
from pyobs.modules import timeout
from .com import com_device


log = logging.getLogger(__name__)


class AscomFocuser(PyObsModule, IFocuser, IFitsHeaderProvider):
    def __init__(self, device: str = None, *args, **kwargs):
        PyObsModule.__init__(self, *args, **kwargs)

        # variables
        self._device = device

    def open(self):
        """Open module."""
        PyObsModule.open(self)

        # init COM
        pythoncom.CoInitialize()

        # do we need to chose a device?
        if not self._device:
            x = win32com.client.Dispatch("ASCOM.Utilities.Chooser")
            x.DeviceType = 'Focuser'
            self._device = x.Choose(None)
            log.info('Selected focuser "%s".', self._device)

        # open connection
        device = win32com.client.Dispatch(self._device)
        if device.Connected:
            log.info('Focuser was already connected.')
        else:
            device.Connected = True
            if device.Connected:
                log.info('Connected to focuser.')
            else:
                raise ValueError('Unable to connect to focuser.')

        # finish COM
        pythoncom.CoInitialize()

    def close(self):
        """Close module."""
        PyObsModule.close(self)

        # get device
        with com_device(self._device) as device:
            # close connection
            if device.Connected:
                log.info('Disconnecting from focuser...')
                device.Connected = False

    def get_fits_headers(self, *args, **kwargs) -> dict:
        """Returns FITS header for the current status of the telescope.

        Returns:
            Dictionary containing FITS headers.
        """

        # get device
        with com_device(self._device) as device:
            # return header
            return {
                'TEL-FOCU': (device.Position / device.StepSize, 'Focus of telescope [mm]')
            }

    @timeout(60000)
    def set_focus(self, focus: float, *args, **kwargs):
        """Sets new focus.

        Args:
            focus: New focus value.
        """

        # get device
        with com_device(self._device) as device:
            # calculating new focus and move it
            log.info('Moving focus to %.2fmm...', focus)
            foc = int(focus * device.StepSize)
            device.Move(foc)

            # wait for it
            while device.IsMoving:
                time.sleep(0.1)

            # finished
            log.info('Reached new focus of %.2mm.', device.Position / device.StepSize)

    def get_focus(self, *args, **kwargs) -> float:
        """Return current focus.

        Returns:
            Current focus.
        """

        # get device
        with com_device(self._device) as device:
            # return current focus
            return device.Position / device.StepSize

    def get_motion_status(self, device: str = None) -> IMotion.Status:
        pass


__all__ = ['AscomFocuser']

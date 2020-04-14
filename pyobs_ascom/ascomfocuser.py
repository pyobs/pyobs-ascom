import logging
import threading
import time
import pythoncom
import win32com.client

from pyobs import PyObsModule
from pyobs.interfaces import IFocuser, IFitsHeaderProvider, IMotion
from pyobs.mixins import MotionStatusMixin
from pyobs.modules import timeout
from pyobs.utils.threads import LockWithAbort
from .com import com_device


log = logging.getLogger(__name__)


class AscomFocuser(MotionStatusMixin, IFocuser, IFitsHeaderProvider, PyObsModule):
    def __init__(self, device: str = None, *args, **kwargs):
        PyObsModule.__init__(self, *args, **kwargs)

        # variables
        self._device = device
        self._focus_offset = 0

        # allow to abort motion
        self._lock_motion = threading.Lock()
        self._abort_motion = threading.Event()

        # init mixins
        MotionStatusMixin.__init__(self, motion_status_interfaces=['IFocuser'])

    def open(self):
        """Open module."""
        PyObsModule.open(self)

        # do we need to chose a device?
        if not self._device:
            # init COM
            pythoncom.CoInitialize()

            # ask user
            x = win32com.client.Dispatch("ASCOM.Utilities.Chooser")
            x.DeviceType = 'Focuser'
            self._device = x.Choose(None)
            log.info('Selected focuser "%s".', self._device)

            # finish COM
            pythoncom.CoInitialize()

        # open mixins
        MotionStatusMixin.open(self)

    def init(self, *args, **kwargs):
        """Initialize device.

        Raises:
            ValueError: If device could not be initialized.
        """
        pass

    def park(self, *args, **kwargs):
        """Park device.

        Raises:
            ValueError: If device could not be parked.
        """
        pass

    def get_fits_headers(self, namespaces: list = None, *args, **kwargs) -> dict:
        """Returns FITS header for the current status of this module.

        Args:
            namespaces: If given, only return FITS headers for the given namespaces.

        Returns:
            Dictionary containing FITS headers.
        """

        # get device
        with com_device(self._device) as device:
            # return header
            # StepSize is in microns, so divide by 1000
            return {
                'TEL-FOCU': (device.Position / device.StepSize / 1000., 'Focus of telescope [mm]')
            }

    @timeout(60000)
    def set_focus(self, focus: float, *args, **kwargs):
        """Sets new focus.

        Args:
            focus: New focus value.
        """

        # set focus + offset
        self._set_focus(focus + self._focus_offset)

    def set_focus_offset(self, offset: float, *args, **kwargs):
        """Sets focus offset.

        Args:
            offset: New focus offset.

        Raises:
            InterruptedError: If focus was interrupted.
        """

        # get current focus (without offset)
        focus = self.get_focus()

        # set offset
        self._focus_offset = offset

        # go to focus
        self._set_focus(focus + self._focus_offset)

    def _set_focus(self, focus):
        """Actually sets new focus.

        Args:
            focus: New focus value.
        """

        # acquire lock
        with LockWithAbort(self._lock_motion, self._abort_motion):
            # get device
            with com_device(self._device) as device:
                # calculating new focus and move it
                log.info('Moving focus to %.2fmm...', focus)
                self._change_motion_status(IMotion.Status.SLEWING, interface='IFocuser')
                foc = int(focus * device.StepSize * 1000.)
                device.Move(foc)

                # wait for it
                while abs(device.Position - foc) > 10:
                    # abort?
                    if self._abort_motion.is_set():
                        log.warning('Setting focus aborted.')
                        return

                    # sleep a little
                    time.sleep(0.1)

                # finished
                log.info('Reached new focus of %.2fmm.', device.Position / device.StepSize / 1000.)
                self._change_motion_status(IMotion.Status.POSITIONED, interface='IFocuser')

    def get_focus(self, *args, **kwargs) -> float:
        """Return current focus.

        Returns:
            Current focus.
        """

        # get device
        with com_device(self._device) as device:
            # return current focus - offset
            return device.Position / device.StepSize / 1000. - self._focus_offset

    def get_focus_offset(self, *args, **kwargs) -> float:
        """Return current focus offset.

        Returns:
            Current focus offset.
        """
        return self._focus_offset

    def stop_motion(self, device: str = None, *args, **kwargs):
        """Stop the motion.

        Args:
            device: Name of device to stop, or None for all.
        """

        # get device
        with com_device(self._device) as device:
            # stop motion
            return device.Halt()

    def is_ready(self, *args, **kwargs) -> bool:
        """Returns the device is "ready", whatever that means for the specific device.

        Returns:
            True, if telescope is initialized and not in an error state.
        """
        return True


__all__ = ['AscomFocuser']

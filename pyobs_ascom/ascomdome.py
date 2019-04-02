import logging
import pythoncom
import win32com.client

from pyobs.interfaces import IMotion
from pyobs.modules import timeout
from pyobs.modules.roof import BaseRoof
from .com import  com_device


log = logging.getLogger('pyobs')


class AscomDome(BaseRoof):
    def __init__(self, device: str = None, *args, **kwargs):
        BaseRoof.__init__(self, *args, **kwargs)

        # variables
        self._device = device
        self._motion_status = IMotion.Status.IDLE

    def open(self):
        """Open module.

        Raises:
            ValueError: If cannot connect to device.
        """
        BaseRoof.open(self)

        # init COM
        pythoncom.CoInitialize()

        # do we need to chose a device?
        if not self._device:
            x = win32com.client.Dispatch("ASCOM.Utilities.Chooser")
            x.DeviceType = 'Dome'
            self._device = x.Choose(None)
            log.info('Selected dome "%s".', self._device)

        # open connection
        device = win32com.client.Dispatch(self._device)
        if device.Connected:
            log.info('Dome was already connected.')
        else:
            device.Connected = True
            if device.Connected:
                log.info('Connected to dome.')
            else:
                raise ValueError('Unable to connect to dome.')

        # finish COM
        pythoncom.CoInitialize()

    def close(self):
        """Clode module."""
        BaseRoof.close(self)

        # get device
        with com_device(self._device) as device:
            # close connection
            if device.Connected:
                log.info('Disconnecting from dome...')
                device.Connected = False

    @timeout(60000)
    def open_roof(self, *args, **kwargs):
        """Open the roof."""

        # get device
        with com_device(self._device) as device:
            # send event
            self._change_motion_status(IMotion.Status.INITIALIZING)

            # open
            device.OpenShutter()

            # send event
            self._change_motion_status(IMotion.Status.IDLE)

    @timeout(60000)
    def close_roof(self, *args, **kwargs):
        """Close the roof."""

        # get device
        with com_device(self._device) as device:
            # send event
            self._change_motion_status(IMotion.Status.PARKING)

            # close
            device.CloseShutter()

            # send event
            self._change_motion_status(IMotion.Status.PARKED)

    def get_percent_open(self) -> float:
        """Get the percentage the roof is open."""

        # not supported
        return 0

    def stop_motion(self, device: str = None):
        """Stop the motion.

        Args:
            device: Name of device to stop, or None for all.
        """

        # not supported
        pass


__all__ = ['AscomDome']

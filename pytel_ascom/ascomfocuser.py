import logging
import time
import pythoncom
import win32com.client

from pytel import PytelModule
from pytel.interfaces import IFocuser, IFitsHeaderProvider, IStatus
from pytel.modules import timeout


log = logging.getLogger('pytel')


class AscomFocuser(PytelModule, IFocuser, IStatus, IFitsHeaderProvider):
    def __init__(self, *args, **kwargs):
        PytelModule.__init__(self, *args, **kwargs)

        # variables
        self._focuser = None

    def open(self):
        # init COM
        pythoncom.CoInitialize()

        # init focuser
        self._focuser = win32com.client.Dispatch(self.config['device'])
        if self._focuser.Connected:
            log.info('Focuser was already connected.')
        else:
            self._focuser.Connected = True
            if self._focuser.Connected:
                log.info('Connected to focuser.')
            else:
                log.info('Unable to connect to focuser.')
                raise ValueError('Could not connect to focuser.')

    def close(self):
        # close connection
        if self._focuser.Connected:
            log.info('Disconnecting from focuser...')
            self._focuser.Connected = False

    @classmethod
    def default_config(cls):
        cfg = super(AscomFocuser, cls).default_config()
        cfg['device'] = None
        return cfg

    def status(self, *args, **kwargs) -> dict:
        """returns current status"""
        # init COM in thread
        pythoncom.CoInitialize()

        # get status
        s = super().status(*args, **kwargs)

        # focus
        if self._focuser.Connected:
            s['focus'] = self._focuser.Position / self._focuser.StepSize

        # finished
        return s

    def get_fits_headers(self, *args, **kwargs) -> dict:
        """get FITS header for the saved status of the telescope"""
        # init COM in thread
        pythoncom.CoInitialize()

        # return header
        return {
            'TEL-FOCU': (self._focuser.Position / self._focuser.StepSize, 'Focus of telescope [mm]')
        }

    @timeout(60000)
    def set_focus(self, focus: float, *args, **kwargs) -> bool:
        """sets focus"""
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
        return True

    def get_focus(self, *args, **kwargs) -> float:
        """returns focus"""
        # init COM in thread
        pythoncom.CoInitialize()

        # return current focus
        return self._focuser.Position / self._focuser.StepSize


__all__ = ['AscomFocuser']

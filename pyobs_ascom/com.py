import pythoncom


@contextmanager
def com_device(device):
    # init COM
    pythoncom.CoInitialize()

    try:
        # dispatch COM object
        yield win32com.client.Dispatch(device)
    finally:
        # finish COM
        pythoncom.CoUninitialize()

import win32com.client as win32


def runFusion():
    excel = win32.gencache.EnsureDispatch('Excel.Application')

    excel.Application.Quit()

if __name__ == '__main__':
    runFusion()
import logging
import win32gui, win32com.client, win32con
import datetime 
import traceback 

class MyWindow():
    def __init__(self, windowId: int, windowName: str) -> None:
        self.id = windowId
        self.title = windowName

    def __repr__(self):
        return f'\n"{self.title}" ({self.id})'

    # def activate(self):
    #     logging.debug(f'activateWindow {self.title}')
    #     try:
    #         win32gui.ShowWindow(self.id, win32con.SW_MAXIMIZE)
    #     except: 
    #         pass
    #     try: 
    #         windowHandel = win32gui.FindWindow(None, self.title)

    #         shell = win32com.client.Dispatch("WScript.Shell")
    #         shell.SendKeys('%')
    #         win32gui.SetForegroundWindow(windowHandel)
    #         t = datetime.datetime.now().strftime("%H:%M:%S")
    #         print(f'{t} -- Activated "{self.title}"')
    #     except: 
    #         logging.debug(f'Window activation failed for "{self.title}" -- {traceback.format_exc()}')
    #     return True

if __name__ == '__main__':
    w1 = MyWindow(0,"bla1")
    w2 = MyWindow(1,"bla2")
    windows = [w1,w2]
    print(windows)
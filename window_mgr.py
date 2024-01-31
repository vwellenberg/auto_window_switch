from time import sleep
from window import *
import config
import win32gui


# TODO function descr
# TODO param checks 
# TODO funcs: configure, pause  via key press listener or input (?) -> threads (?)

def winEnumHandler( hwnd, ctx ):
    if win32gui.IsWindowVisible( hwnd ):
        title = win32gui.GetWindowText( hwnd )
        if title:
            # print(f"{win32gui.GetWindowText( hwnd )}", f'({hwnd})')
            config.windows.append(MyWindow(windowId=hwnd,windowName=win32gui.GetWindowText( hwnd )))

# TODO thread ?
class WindowMgr():

    # TODO possibility to use selection 
    def __init__(self, interval:int = 30, game_selection:list = []) -> None:
        self.__interval = interval
        self.load_windows()
        self.windows = config.windows
        self.check()
        
        print(self.windows)
        # self.filter_windows()
        # print(self.windows)

        # self.select_windows(windows)
        # self.start()

    def check(self):
        # TODO impl
        if not len(config.windows):
            # TODO error
            print("NO WINDOWS")
        pass

    def load_windows(self):
        win32gui.EnumWindows( winEnumHandler, None )

    def filter_windows(self):
        windows = self.windows
        filter = ["Calculator", "NVIDIA", "Program Manager", "Settings"]
        for window in windows: 
            if filter in window:
                windows.remove(window)
        self.windows = windows

    def select_windows(self,windows):
        print("select")
        idx = 0
        for idx in range (0,len(windows)):
            print(idx, f'{windows[idx].title} ({windows[idx].id})')

    # def focusNextWindow(self):
    #     try:
    #         window = self.__GAMES[self.idx]
    #         # TODO timestamp 
    #         print("new window",self.idx, window.windowName, window.windowId)
    #         windowHandel = win32gui.FindWindow(None, window.windowName)
    #         win32gui.SetForegroundWindow(windowHandel)
    #         self.idx += 1
    #         if self.idx > len(self.__GAMES) - 1:
    #             self.idx = 0
    #     # TODO more error info
    #     except Exception as ex: print('Error', ex)

    # def start(self):
    #     print()
    #     # TODO print chosen windows
    #     input('go?')
    #     do_run = True
    #     while(do_run):
    #         sleep(self.__interval)
    #         self.focusNextWindow()
            
            
if __name__ == '__main__':

    ws = WindowMgr()
    

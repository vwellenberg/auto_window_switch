from time import sleep
from window import *
import win32gui
import logging
import yaml 
import datetime 
import os 
import win32con

# TODO (!!!) impl function to remove windows from queue
# TODO docstrings
# TODO pause via key press listener or input (?) -> threads (?)

raw_windows = []
"""Needed as a container for the initial reading of the current windows."""

def winEnumHandler( hwnd, ctx ):
    if win32gui.IsWindowVisible( hwnd ):
        title = win32gui.GetWindowText( hwnd )
        if title:
            # logging.debug(f"{win32gui.GetWindowText( hwnd )}", f'({hwnd})')
            raw_windows.append(MyWindow(windowId=hwnd,windowName=win32gui.GetWindowText( hwnd )))

# TODO thread ?
class WindowMgr():

    def __init__(self, interval:int = 60, selection:list = []) -> None:
        # TODO validate
        interval = int(input('Window focus duration (in seconds): ')) 
        # interval = 3 # for debug
        self.idx_last_window = -1
        self.idx_active_window = 0
        self.countdown_s = 3
        self.__interval = interval
        self.load_windows()
        self.windows = raw_windows # save windows as attribute
        self.selected_windows = selection
        """windows chosen to cycle through"""
        self.validate()

        logging.debug(f'raw windows: {self.windows}')
        # TODO Fix: Has to be called more than once atm
        self.prefilter_windows()
        self.prefilter_windows()
        logging.debug(f'filtered windows: {self.windows}')

        self.prompt_window_selection()
        self.start()

    def validate(self):
        # TODO impl
        if not len(raw_windows):
            # TODO error
            logging.error("No Windows")

    def load_windows(self):
        win32gui.EnumWindows( winEnumHandler, None )

    def read_filters(self):
        with open('filters.yaml','r') as f:
            filters = yaml.safe_load(f)
        self.filters = filters
        logging.debug(self.filters)

    def prefilter_windows(self):
        filtered = 0
        windows = self.windows
        self.read_filters()

        for filter in self.filters:
            for window in windows:
                if filter in window.title:
                    windows.remove(window)
                    logging.debug(f'Removed: "{window.title} ({window.id})" by Filter: "{filter}"')
                    filtered += 1
        self.windows = windows
        logging.debug(f'number of filtered windows: {filtered}')

    def print_selectable_windows(self):
        print('Selectable windows:')
        for idx in range (0,len(self.windows)):
            print(f'{idx} - "{self.windows[idx].title}"')
        print()

    def prompt_window_selection(self):
        choose_more = True
        while choose_more:
            os.system('cls') # clear screen

            self.print_selectable_windows()

            if self.selected_windows:
                self.print_selected_windows()

            print()
            i = input('Add new game by index (or press enter to finish): ')
            # check input
            try: 
                i = int(i)
                idx = i
                if idx > len(self.windows) -1:
                    # invalid
                    os.system('cls') # clear screen
                    input('<Invalid index number> (press enter to continue)')
                else: 
                    # valid input, add
                    self.selected_windows.append(self.windows[idx])
                    print(f'"{self.windows[idx].title}" added')
            except:
                # no integer
                # check other inputs than index number
                if i == '' or i == 'exit':
                    choose_more = False
                else: 
                    os.system('cls') # clear screen
                    input('<Invalid input> (press enter to continue)')

    def activateWindow(self,idx:int):
        window = self.selected_windows[idx]
        logging.debug(f'activateWindow {window.title}')
        win32gui.ShowWindow(window.id, win32con.SW_MAXIMIZE)
        windowHandel = win32gui.FindWindow(None, window.title)
        win32gui.SetForegroundWindow(windowHandel)

    def nextWindow(self):
        success = False
        while (not success):
            try:
                self.activateWindow(self.idx_active_window)

                self.idx_active_window += 1
                if self.idx_active_window > len(self.selected_windows) - 1:
                    self.idx_active_window = 0

                self.idx_last_window += 1
                if self.idx_last_window > len(self.selected_windows) - 1:
                    self.idx_last_window = 0

                success = True
            except Exception: 
                try:
                    logging.debug('Switch did not work ... trying again!')
                    sleep(0.1)
                except Exception:
                    pass
        t = datetime.datetime.now().strftime("%H:%M:%S")
        print(f'{t} -- Activated "{self.selected_windows[self.idx_last_window].title}"')

    def print_selected_windows(self):
        print('Chosen window order so far:')
        for idx in range(0,len(self.selected_windows)):
            print(f'{idx+1}. {self.selected_windows[idx].title}')
        print()

    def countdown(self):
        while (self.countdown_s > 0):
            os.system('cls') # clear screen
            self.print_selected_windows()
            print()
            print(f'Starting in {self.countdown_s} ...')
            sleep(1)
            self.countdown_s -= 1

    def start(self):
        do_run = True
        os.system('cls') # clear screen

        input('Go? (Press enter to continue)')

        self.countdown()
        
        os.system('cls')
        while(do_run):
            self.nextWindow()
            sleep(self.__interval)

if __name__ == '__main__':
    logging.basicConfig(
        filename='logfile.log',
        filemode='w',
        encoding='utf-8',
        format='%(asctime)s %(levelname)-8s %(message)s',
        level=logging.DEBUG,
        datefmt='%Y-%m-%d %H:%M:%S')

    ws = WindowMgr()
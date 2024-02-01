from time import sleep
from window import *
import win32gui, win32com.client
import logging
import yaml 
import datetime 
import os 
import win32con
import traceback
import sys

# TODO (!!!) impl function to remove windows from queue
# TODO docstrings
# TODO pause via key press listener or input (?) -> threads (?)

windows = []
"""Needed as a container for the initial reading of the current windows."""

def winEnumHandler( hwnd, ctx ):
    if win32gui.IsWindowVisible( hwnd ):
        title = win32gui.GetWindowText( hwnd )
        if title:
            # logging.debug(f"{win32gui.GetWindowText( hwnd )}", f'({hwnd})')
            windows.append(MyWindow(windowId=hwnd,windowName=win32gui.GetWindowText( hwnd )))

# TODO thread ?
class WindowMgr():

    def __init__(self, interval:int = 60, selection:list = []) -> None:
        # TODO validate
        interval = int(input('Window focus duration (in seconds): ')) 
        # interval = 3 # for debug
        self.__idx_last_window = -1
        self.__idx_next_window = 0
        self.__countdown_s = 3
        self.__interval = interval
        self.init_windows()
        self.__windows = windows # save windows as attribute
        self.__window_queue = selection
        """windows chosen to cycle through"""
        self.validate()

        logging.debug(f'raw windows: {self.__windows}')
        # TODO Fix: Has to be called more than once atm
        self.prefilter_windows()
        self.prefilter_windows()
        logging.debug(f'filtered windows: {self.__windows}')

        self.prompt_window_selection()
        self.start()

    def validate(self):
        # TODO impl
        if not len(windows):
            # TODO error
            logging.error("No Windows")

    def reload_windows(self):
        self.init_windows()
        self.__windows = windows

    def init_windows(self):
        """Loads all existing windows"""
        win32gui.EnumWindows( winEnumHandler, None )

    def read_filters(self):
        with open('filters.yaml','r') as f:
            filters = yaml.safe_load(f)
        self.filters = filters
        logging.debug(self.filters)

    def prefilter_windows(self):
        filtered = 0
        windows = self.__windows
        self.read_filters()

        for filter in self.filters:
            for window in windows:
                if filter in window.title:
                    windows.remove(window)
                    logging.debug(f'Removed: "{window.title} ({window.id})" by Filter: "{filter}"')
                    filtered += 1
        self.__windows = windows
        logging.debug(f'number of filtered windows: {filtered}')

    def print_selectable_windows(self):
        print('Selectable windows:')
        for idx in range (0,len(self.__windows)):
            print(f'{idx} - "{self.__windows[idx].title}"')
        print()

    def prompt_window_selection(self):
        choose_more = True
        while choose_more:
            os.system('cls') # clear screen

            self.print_selectable_windows()

            if self.__window_queue:
                self.print_window_queue()

            print()
            i = input('Add new game by index (or press enter to finish): ')
            # check input
            try: 
                i = int(i)
                idx = i
                if idx > len(self.__windows) -1:
                    # invalid
                    os.system('cls') # clear screen
                    input('<Invalid index number> (press enter to continue)')
                else: 
                    # valid input, add
                    self.__window_queue.append(self.__windows[idx])
                    print(f'"{self.__windows[idx].title}" added')
            except:
                # no integer
                # check other inputs than index number
                if i == '' or i == 'exit':
                    choose_more = False
                else: 
                    os.system('cls') # clear screen
                    input('<Invalid input> (press enter to continue)')

    def activate_window(self,idx:int):
        window = self.__window_queue[idx]
        logging.debug(f'activateWindow {window.title}')
        try:
            win32gui.ShowWindow(window.id, win32con.SW_MAXIMIZE)
        except: 
            pass
        try: 
            windowHandel = win32gui.FindWindow(None, window.title)

            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys('%')
            win32gui.SetForegroundWindow(windowHandel)
        except: 
            print(traceback.format_exc())

    def does_window_exist(self,title):
        exists = False
        for idx in range (0,len(self.__windows)):
            if title == self.__windows[idx].title:
                exists = True
        if not exists:
            logging.debug(f'"{title}" does not exist anymore')
        return exists

    def update_window_pointers(self):
        self.__idx_next_window += 1
        if self.__idx_next_window > len(self.__window_queue) - 1:
            self.__idx_next_window = 0

        self.__idx_last_window += 1
        if self.__idx_last_window > len(self.__window_queue) - 1:
            self.__idx_last_window = 0

    def handle_closed_window(self):
        """Check for window existence. If not, remove it from the queue"""
        self.reload_windows() # check currently active windows
        __next_window = self.__window_queue[self.__idx_next_window]
        if not self.does_window_exist(__next_window.title):
            print(f'{__next_window.title} does not exist anymore')
            self.__window_queue.remove(self.__idx_next_window)
            logging.debug(f'"{__next_window.title()}" does not exist anymore')
            # next window
            self.update_window_pointers()
        else:
            print("window exists")

    def next_window(self):
        success = False

        # self.checkWindows()

        while (not success):
            try:
                self.handle_closed_window()
                self.activate_window(self.__idx_next_window)
                self.update_window_pointers()

                success = True
            except Exception: 
                try:
                    print(traceback.format_exc())
                    logging.debug('Switch did not work ... trying again!')
                    sleep(0.1)
                except Exception:
                    pass
        t = datetime.datetime.now().strftime("%H:%M:%S")
        print(f'{t} -- Activated "{self.__window_queue[self.__idx_last_window].title}"')

    def print_window_queue(self):
        print('Chosen window order so far:')
        for idx in range(0,len(self.__window_queue)):
            print(f'{idx+1}. {self.__window_queue[idx].title}')
        print()

    def countdown(self):
        while (self.__countdown_s > 0):
            os.system('cls') # clear screen
            self.print_window_queue()
            print()
            print(f'Starting in {self.__countdown_s} ...')
            sleep(1)
            self.__countdown_s -= 1

    def start(self):
        do_run = True
        os.system('cls') # clear screen

        input('Go? (Press enter to continue)')

        self.countdown()
        
        os.system('cls')
        while(do_run):
            self.next_window()
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
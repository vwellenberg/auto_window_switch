from time import sleep
from my_window import *
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

def clear_screen():
    os.system('cls')

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
        interval = 60
        try:
            interval = int(input('Input window switch interval (in seconds) or press enter for 60 seconds: '))
        except:
            interval = 60
        self._MAX_RETRY = 50
        self.cycles = 0
        self._idx_last_window = -1
        self._idx_next_window = 0
        self._countdown_s = 3
        self._interval = interval
        self.init_windows()
        self._found_windows = windows # save windows as attribute
        """windows found on ... Windows"""
        self._window_queue = selection
        """windows chosen to cycle through"""
        self.validate()

        logging.debug(f'raw windows: {self._found_windows}')
        # TODO Fix: Has to be called more than once atm
        self.prefilter_windows()
        self.prefilter_windows()
        logging.debug(f'filtered windows: {self._found_windows}')

        self.prompt_window_selection()
        self.start()

    def validate(self):
        # TODO impl
        if not len(windows):
            # TODO error
            logging.error("No Windows")

    # def reload_windows(self):
    #     # reset loaded Windows 
    #     windows = []
    #     self.init_windows()
    #     # self._found_windows = windows
    #     logging.debug(f'reloaded windows: {windows}')

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
        windows = self._found_windows
        self.read_filters()

        for filter in self.filters:
            for window in windows:
                if filter in window.title:
                    windows.remove(window)
                    logging.debug(f'Removed: "{window.title} ({window.id})" by Filter: "{filter}"')
                    filtered += 1
        self._found_windows = windows
        logging.debug(f'number of filtered windows: {filtered}')

    def print_selectable_windows(self):
        print('Selectable windows:')
        for idx in range (0,len(self._found_windows)):
            print(f'{idx} - "{self._found_windows[idx].title}"')
        print()

    def prompt_window_selection(self):
        choose_more = True
        while choose_more:
            clear_screen()

            self.print_selectable_windows()

            if self._window_queue:
                self.print_window_queue()

            print()
            i = input('Add new game by index (or press enter to finish): ')
            # check input
            try: 
                i = int(i)
                idx = i
                if idx > len(self._found_windows) -1:
                    # invalid
                    clear_screen() 
                    input('<Invalid index number> (press enter to continue)')
                else: 
                    # valid input, add
                    self._window_queue.append(self._found_windows[idx])
                    print(f'"{self._found_windows[idx].title}" added')
            except:
                # no integer
                # check other inputs than index number
                if i == '' or i == 'exit':
                    choose_more = False
                else: 
                    clear_screen() 
                    input('<Invalid input> (press enter to continue)')

    def _is_window_existent(self,window):
        # TODO Fix
        exists = False
        # update active windows from ... Windows
        # self.reload_windows()
        # iterate over windows
        for idx in range (0,len(self._found_windows)):
            if window.title == self._found_windows[idx].title:
                logging.debug(f'"{window.title}" found ... match with "{self._found_windows[idx].title}"')
                exists = True
                break
        if not exists:
            logging.debug(f'Did not find "{window.title}"')
        return exists

    def _update_window_pointers(self):
        self._idx_next_window += 1
        if self._idx_next_window > len(self._window_queue) - 1:
            self._idx_next_window = 0

        self._idx_last_window += 1
        if self._idx_last_window > len(self._window_queue) - 1:
            self._idx_last_window = 0

    def activate_window(self,idx:int) -> bool:
        """core functionality to get a window to the foreground"""
        window = self._window_queue[idx]
        success = False
        logging.debug(f'activateWindow {window.title}')
        try:
            win32gui.ShowWindow(window.id, win32con.SW_MAXIMIZE)
        except: 
            pass
        try: 
            windowHandel = win32gui.FindWindow(None, window.title)

            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys('%') # Alt-key should improve reliability
            shell.SendKeys('%')
            win32gui.SetForegroundWindow(windowHandel)
            self._update_window_pointers()
            t = datetime.datetime.now().strftime("%H:%M:%S")
            print(f'{t} -- Activated "{self._window_queue[self._idx_last_window].title}"')
            success = True
        except: 
            logging.debug(f'Window activation failed for "{window.title}" -- {traceback.format_exc()} ... trying again!')
            sleep(0.1)
        return success
    
    def _cycle(self):
        window = self._window_queue[self._idx_next_window]
        if not self._is_window_existent(window):   
            # window does not exist, remove it from the queue
            self._window_queue.remove(self._idx_next_window) # TODO Fix
            logging.debug(f'"{window.title}" does not exist and got removed from queue')
            # go to next title
            self._update_window_pointers()
        success = False
        attempts = 0
        while (not success and attempts < self._MAX_RETRY):
            logging.debug(f'Attempt #{attempts} to activate "{window.title}"')
            # try to 'activate' window as often as possible
            success = self.activate_window(self._idx_next_window)
            attempts += 1
        if not success:
            t = datetime.datetime.now().strftime("%H:%M:%S")
            print(f'{t} -- Activation failed for "{self._window_queue[self._idx_last_window].title}" after {self._MAX_RETRY} attempts')
        self.cycles += 1

    def print_window_queue(self):
        print('Current window order:')
        for idx in range(0,len(self._window_queue)):
            print(f'{idx+1}. {self._window_queue[idx].title}')
        print()

    def countdown(self):
        while (self._countdown_s > 0):
            clear_screen() 
            self.print_window_queue()
            print()
            print(f'Starting in {self._countdown_s} ...')
            sleep(1)
            self._countdown_s -= 1

    def start(self):
        do_run = True
        clear_screen()

        input('Start? (Press enter to continue)')

        self.countdown()
        
        clear_screen()
        while(do_run):
            self._cycle()
            sleep(self._interval)

if __name__ == '__main__':
    logging.basicConfig(
        filename='logfile.log',
        filemode='w',
        encoding='utf-8',
        format='%(asctime)s %(levelname)-8s %(message)s',
        level=logging.DEBUG,
        datefmt='%Y-%m-%d %H:%M:%S')

    ws = WindowMgr()
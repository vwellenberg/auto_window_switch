from time import sleep
from my_window import MyWindow
import win32gui
import win32com.client
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
    """Clear terminal text"""
    os.system('cls')


def winEnumHandler(hwnd, ctx):
    if win32gui.IsWindowVisible(hwnd):
        title = win32gui.GetWindowText(hwnd)
        if title:
            # logging.debug(f"{win32gui.GetWindowText( hwnd )}", f'({hwnd})')
            windows.append(
                MyWindow(
                    windowId=hwnd,
                    windowName=win32gui.GetWindowText(hwnd)))

# TODO thread ?


class WindowMgr():

    def __init__(self, interval: int = 60, selection: list = []) -> None:
        interval = 60
        try:
            user_info = 'Input window switch interval (in seconds) or press enter for 60 seconds: '
            interval = int(input(user_info))
        except BaseException:
            interval = 60
        self._retry_interval = 1.0
        self._max_retry = 10
        self.cycles = 0
        self._idx_last_window = -1
        self._idx_next_window = 0
        self._countdown_s = 3
        self._interval = interval
        self._init_windows()
        self._found_windows = windows  # save windows as attribute
        """windows found on ... Windows"""
        self._window_queue = selection
        """windows chosen to cycle through"""
        self._validate()

        self._read_filters()

        logging.debug('Windows found: %s', self._found_windows)
        # TODO Fix: Has to be called more than once atm
        self._prefilter_windows()
        self._prefilter_windows()
        logging.debug('Filtered windows: %s', self._found_windows)

        self._prompt_window_selection()
        self.start()

    def _validate(self):
        """Check all variables for validity"""
        logging.debug('Validating ...')
        if self._interval < 5:
            logging.error('Invalid input: "_interval" has to be higher than 5')
            input('ERROR -- invalid input: Interval has to be higher than 5')
            exit()
        # # TODO impl
        # if not len(windows):
        #     # TODO error
        #     logging.error("No Windows")

    def _init_windows(self):
        """Loads all existing windows"""
        win32gui.EnumWindows(winEnumHandler, None)

    def _read_filters(self):
        try:
            with open('filters.yaml', 'r', encoding='utf-8') as f:
                filters = yaml.safe_load(f)
                logging.debug('filters.yaml loaded.')
                self.filters = filters
                logging.debug('filters initialized')
        except FileNotFoundError:
            input('WARNING -- "filters.yaml" not found. It needs to be in the same folder as aws.exe. Using no window filters.')
            logging.warn(f'"filters.yaml" not found: {traceback.format_exc()}')
            self.filters = []
        except BaseException:
            input('WARNING -- Reading "filters.yaml" failed. Using no window filters.')
            logging.warn('Reading the "filters.yaml" failed: %s', traceback.format_exc())
            self.filters = []
        logging.debug(self.filters)

    def _prefilter_windows(self):
        filtered = 0
        windows = self._found_windows
        for filter in self.filters:
            for window in windows:
                if filter in window.title:
                    windows.remove(window)
                    logging.debug(
                        f'Removed: "{window.title} ({window.id})" by Filter: "{filter}"')
                    filtered += 1
        self._found_windows = windows
        logging.debug(f'number of filtered windows: {filtered}')

    def _print_selectable_windows(self):
        print('Selectable windows:')
        for idx in range(0, len(self._found_windows)):
            print(f'{idx} - "{self._found_windows[idx].title}"')
        print()

    def _prompt_window_selection(self):
        choose_more = True
        while choose_more:
            clear_screen()

            self._print_selectable_windows()

            if self._window_queue:
                self._print_window_queue()

            print()
            user_input = input('Add new game by index (or press enter to finish): ')
            # check input
            try:
                user_input = int(user_input)
                idx = user_input
                if idx > len(self._found_windows) - 1:
                    # invalid
                    clear_screen()
                    input('<Invalid index number> (press enter to continue)')
                else:
                    # valid input, add
                    self._window_queue.append(self._found_windows[idx])
                    print(f'"{self._found_windows[idx].title}" added')
            except BaseException:
                # no integer
                # check other inputs than index number
                if (user_input == '' or user_input == 'exit') and len(self._window_queue) >= 2 :
                    choose_more = False
                elif len(self._window_queue) <= 2:
                    clear_screen()
                    input('You need to choose at least two windows (press enter to continue)')
                else:
                    clear_screen()
                    input('<Invalid input> (press enter to continue)')

    def _is_window_existent(self, window):
        # TODO Fix
        exists = False
        # update active windows from ... Windows
        # self.reload_windows()
        # iterate over windows
        for idx in range(0, len(self._found_windows)):
            if window.title == self._found_windows[idx].title:
                logging.debug(
                    f'"{window.title}" found ... match with "{self._found_windows[idx].title}"')
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

    def _activate_window(self, idx: int) -> bool:
        """core functionality to get a window to the foreground"""
        window = self._window_queue[idx]
        success = False
        logging.debug(f'activateWindow {window.title}')
        try:
            win32gui.ShowWindow(window.id, win32con.SW_MAXIMIZE)
        except BaseException:
            pass
        try:
            windowHandel = win32gui.FindWindow(None, window.title)

            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys('%')  # Alt-key should improve reliability
            shell.SendKeys('%')
            win32gui.SetForegroundWindow(windowHandel)
            self._update_window_pointers()
            t = datetime.datetime.now().strftime("%H:%M:%S")
            print(
                f'{t} -- Activated "{self._window_queue[self._idx_last_window].title}"')
            success = True
        except BaseException:
            logging.debug(
                f'Window activation failed for "{window.title}" -- {traceback.format_exc()} ... trying again!')
            sleep(self._retry_interval)
        return success

    def _cycle(self):
        window = self._window_queue[self._idx_next_window]
        if not self._is_window_existent(window):
            # window does not exist, remove it from the queue
            self._window_queue.remove(self._idx_next_window)  # TODO Fix
            logging.debug(
                f'"{window.title}" does not exist and got removed from queue')
            # go to next title
            self._update_window_pointers()
        success = False
        attempts = 1
        while (not success and attempts <= self._max_retry):
            logging.debug(f'Attempt #{attempts} to activate "{window.title}"')
            # try to 'activate' window as often as possible
            success = self._activate_window(self._idx_next_window)
            attempts += 1
        if not success:
            t = datetime.datetime.now().strftime("%H:%M:%S")
            print(
                f'{t} -- Activation failed for "{self._window_queue[self._idx_last_window].title}" after {self._max_retry} attempts')
        self.cycles += 1

    def _print_window_queue(self):
        print('Current window order:')
        for idx in range(0, len(self._window_queue)):
            print(f'{idx+1}. {self._window_queue[idx].title}')
        print()

    def _start_countdown(self):
        while (self._countdown_s > 0):
            clear_screen()
            self._print_window_queue()
            print()
            print(f'Starting in {self._countdown_s} ...')
            sleep(1)
            self._countdown_s -= 1

    def start(self):
        do_run = True
        clear_screen()

        input('Start? (Press enter to continue)')

        self._start_countdown()

        clear_screen()
        while (do_run):
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

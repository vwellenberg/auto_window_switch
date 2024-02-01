from window_mgr import WindowMgr
import logging

# TODO tests

if __name__ == '__main__':
    logging.basicConfig(
        filename='logfile.log',
        filemode='w',
        encoding='utf-8',
        format='%(asctime)s %(levelname)-8s %(message)s',
        level=logging.DEBUG,
        datefmt='%Y-%m-%d %H:%M:%S')
    ws = WindowMgr()

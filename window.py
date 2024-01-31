class MyWindow():
    def __init__(self, windowId: int, windowName: str) -> None:
        self.id = windowId
        self.title = windowName

    def __repr__(self):
        return f'\n{self.title} ({self.id})'
        

if __name__ == '__main__':
    w1 = MyWindow(0,"bla1")
    w2 = MyWindow(1,"bla2")
    windows = [w1,w2]
    print(windows)
from tkinter import Tk, Button, HORIZONTAL
from tkinter import ttk

from fileController import FileController

class Window(Tk):
    def __init__(self):
        super().__init__()

        self.title("Hello World")

        self.button = Button(text="My simple app.", command=self.handle_button_press)
        # self.button.bind("", self.handle_button_press)
        self.button.pack()
        self.progress_var = ttk.Progressbar(orient=HORIZONTAL, length=400, mode="determinate")
        self.progress_var.pack(pady=30)
        btn = Button(text="progress", command=self.progress)
        btn.pack(pady=30)
        btn = Button(text="test", command=self.test)
        btn.pack(pady=30)

    def handle_button_press(self):
        self.destroy()

    
    def progress(self):
        # increment of progress value
        self.progress_var["value"] += 20

    def progress_callback(self, iterable):
        total_count = len(iterable)
        step = 100.0 / total_count
        for (i, elem) in enumerate(iterable):
            self.progress_var["value"] = (i+1) * step
            self.update()
            yield elem

    def test(self):
        FileController.progresstest(self.progress_callback)
        pass

import sys

if __name__ == "__main__":
    if ('--tui' in sys.argv or '-h' in sys.argv or '--help' in sys.argv): 
        import tui
    else:
        from ui.window import Window
        window = Window()
        window.mainloop()
        
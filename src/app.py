import sys

if __name__ == "__main__":
    if ('--help' in sys.argv or '-h' in sys.argv): 
        print("elexam is an utility program to make Excel file processing for e-learning a bit easier.")
        print("Usage: app.py [options]")
        print("\nOptions:")
        print("\t -h, --help\tShow this help")
        print("\t --settings\tShow settings file location")
        print("\t -v, --version\tShow program version")
    elif '--version' in sys.argv or '-v' in sys.argv:
        import version
        print(version.VERSION)
    else:
        # from ui.window import Window
        # window = Window()
        # window.mainloop()
        import tui
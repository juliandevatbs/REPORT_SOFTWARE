from tkinter import Tk
from gui import Gui
import sv_ttk

from scripts.print.print_header_data import print_header_data


import hupper
from gui import main


if __name__ == "__main__":

    reloader = hupper.start_reloader('gui.main')
    main()

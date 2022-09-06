import UserInterface as ui
import tkinter as tk

# close the user interface
def closeWindow():
    global mainRoot
    mainRoot.destroy()

def main():
    global mainRoot
    mainRoot = tk.Tk()
    ui.UI(mainRoot)
    mainRoot.mainloop()


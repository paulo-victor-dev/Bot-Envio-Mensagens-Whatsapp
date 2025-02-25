from Interface import Tela_Principal
import sys


if __name__=='__main__':
    if hasattr(sys, '_MEIPASS'):
        from customtkinter import set_appearance_mode

    app = Tela_Principal()
    app.mainloop()
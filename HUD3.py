#!/usr/bin/python3
import tkinter as tk


class MUB3:
    def __init__(self, master=None):
        # build ui
        self.cont=0
        toplevel1 = tk.Tk() if master is None else tk.Toplevel(master)
        toplevel1.configure(
            background="#706665",
            borderwidth=5,
            height=215,
            relief="raised",
            width=415)
        frame2 = tk.Frame(toplevel1)
        frame2.configure(background="#F5F5F5", height=200, width=400)
        self.button2 = tk.Button(frame2)
        self.button2.configure(font="{ARIAL} 12 {bold}", text='OK')
        self.button2.place(
            anchor="nw",
            height=30,
            relx=0.3,
            rely=0.7,
            width=150,
            x=0,
            y=0)
        self.button2.configure(command=self.fecha)
        label1 = tk.Label(frame2)
        label1.configure(
            font="{@Microsoft JhengHei Light} 12 {bold}",
            text='Sucesso! Certificado emitido!',
            foreground="#00B140")
        label1.place(anchor="nw", relx=0.16, rely=0.2, x=0, y=0)
        frame2.pack(side="top")

        # Main widget
        self.mainwindow = toplevel1

    def run(self):
        self.mainwindow.mainloop()
        self.cont+=1

    def fecha(self):
        self.mainwindow.destroy()


def abre():
    app = MUB3()
    app.run()

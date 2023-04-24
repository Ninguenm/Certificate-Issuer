#!/usr/bin/python3
import tkinter as tk
import tkinter.ttk as ttk
import Certificado, HUD2, HUD3
from threading import Thread


class MUB1:
    def __init__(self, master=None):
        self.lis=[]
        toplevel1 = tk.Tk() if master is None else tk.Toplevel(master)
        toplevel1.configure(
            background="#706665",
            borderwidth=5,
            height=715,
            relief="raised",
            width=815)
        frame2 = tk.Frame(toplevel1)
        frame2.configure(background="#F5F5F5", height=700, width=800)
        self.label1 = tk.Label(frame2)
        self.label1.configure(
            background="#BDBDBD",
            font="{ARIAL} 16 {bold italic}",
            text='EMISSÃO DE CERTIFICADO')
        self.label1.place(
            anchor="nw",
            relx=0.28,
            rely=0.04,
            width=300,
            x=0,
            y=0)
        self.entry1 = tk.Entry(frame2)
        self.entry1.place(anchor="nw", relx=0.2, rely=0.30, width=150, x=0, y=0)
        self.label2 = tk.Label(frame2)
        self.label2.configure(
            background="#E0E0E0",
            font="{ARIAL} 10 {bold}",
            relief="ridge",
            text='CNPJ')
        self.label2.place(
            anchor="nw",
            height=20,
            relx=0.05,
            rely=0.30,
            width=100,
            x=0,
            y=0)
        button1 = tk.Button(frame2)
        button1.configure(text='Encontrar Cliente')
        button1.place(
            anchor="nw",
            height=20,
            relx=0.13,
            rely=0.35,
            width=100,
            x=0,
            y=0)
        button1.configure(command = self.pegarCliente)
        self.label5 = tk.Label(frame2)
        self.label5.configure(
            background="#E0E0E0",
            font="{ARIAL} 10 {bold}",
            relief="ridge",
            text='ESTADO')
        self.label5.place(
            anchor="nw",
            height=20,
            relx=0.47,
            rely=0.34,
            width=100,
            x=0,
            y=0)
        self.label4 = tk.Label(frame2)
        self.label4.configure(
            background="#E0E0E0",
            font="{ARIAL} 10 {bold}",
            relief="ridge",
            text='ENDEREÇO')
        self.label4.place(
            anchor="nw",
            height=20,
            relx=0.47,
            rely=0.30,
            width=100,
            x=0,
            y=0)
        self.label3 = tk.Label(frame2)
        self.label3.configure(
            background="#E0E0E0",
            font="{ARIAL} 10 {bold}",
            relief="ridge",
            text='NOME')
        self.label3.place(
            anchor="nw",
            height=20,
            relx=0.47,
            rely=0.26,
            width=100,
            x=0,
            y=0)
        self.label6 = tk.Label(frame2)
        self.label6.configure(
            background="#E0E0E0",
            font="{ARIAL} 20 {bold}",
            relief="ridge")
        self.label6.place(
            anchor="nw",
            height=20,
            relx=0.62,
            rely=0.26,
            width=280,
            x=0,
            y=0)
        self.label7 = tk.Label(frame2)
        self.label7.configure(
            background="#E0E0E0",
            font="{ARIAL} 20 {bold}",
            relief="ridge")
        self.label7.place(
            anchor="nw",
            height=20,
            relx=0.62,
            rely=0.30,
            width=280,
            x=0,
            y=0)
        self.label8 = tk.Label(frame2)
        self.label8.configure(
            background="#E0E0E0",
            font="{ARIAL} 20 {bold}",
            relief="ridge")
        self.label8.place(
            anchor="nw",
            height=20,
            relx=0.62,
            rely=.34,
            width=280,
            x=0,
            y=0)
        button2 = tk.Button(frame2)
        button2.configure(font="{ARIAL} 10 {bold}", text='EMITIR')
        button2.place(
            anchor="nw",
            height=20,
            relx=0.8,
            rely=0.9,
            width=100,
            x=0,
            y=0)
        button2.configure(command=self.gerarLis)

        '''button7 = tk.Button(frame2)
        button7.configure(font="{ARIAL} 10 {bold}", text='ABRIR')
        button7.place(
            anchor="nw",
            height=20,
            relx=0.8,
            rely=0.95,
            width=100,
            x=0,
            y=0)
        button7.configure(command=self.abrirDoc)'''
        
        separator1 = ttk.Separator(frame2)
        separator1.configure(orient="horizontal")
        separator1.place(
            anchor="nw",
            relx=0.06,
            rely=0.4,
            width=700,
            x=0,
            y=0)
        self.entry2 = ttk.Entry(frame2)
        self.entry2.place(anchor="nw", relx=0.44, rely=0.47, x=0, y=0)
        self.label9 = tk.Label(frame2)
        self.label9.configure(
            background="#E0E0E0",
            font="{ARIAL} 10 {bold}",
            relief="ridge",
            text='Quantos itens serão certificados?')
        self.label9.place(
            anchor="nw",
            height=20,
            relx=0.05,
            rely=0.47,
            width=300,
            x=0,
            y=0)
        button3 = tk.Button(frame2)
        button3.configure(text='Confimar')
        button3.place(
            anchor="nw",
            height=20,
            relx=0.37,
            rely=0.51,
            width=100,
            x=0,
            y=0)
        button3.configure(command=self.pegarQTDITENS)
        self.label10 = tk.Label(frame2)
        self.label10.configure(
            background="#E0E0E0",
            font="{ARIAL} 10 {bold}",
            relief="ridge",
            text='X Itens Serão certificados')
        self.label10.place(
            anchor="nw",
            height=20,
            relx=0.7,
            rely=0.42,
            width=200,
            x=0,
            y=0)
        separator3 = ttk.Separator(frame2)
        separator3.configure(orient="vertical")
        separator3.place(
            anchor="nw",
            height=100,
            relx=0.64,
            rely=0.42,
            width=1,
            x=0,
            y=0)
        separator4 = ttk.Separator(frame2)
        separator4.configure(orient="horizontal")
        separator4.place(
            anchor="nw",
            relx=0.06,
            rely=0.58,
            width=700,
            x=0,
            y=0)
        self.label11 = tk.Label(frame2)
        self.label11.configure(
            background="#E0E0E0",
            font="{ARIAL} 10 {bold}",
            relief="ridge",
            text='Código do item')
        self.label11.place(
            anchor="nw",
            height=20,
            relx=0.15,
            rely=0.64,
            width=150,
            x=0,
            y=0)
        self.entry3 = ttk.Entry(frame2)
        self.entry3.place(anchor="nw", relx=0.36, rely=0.64, width=50, x=0, y=0)
        button4 = tk.Button(frame2)
        button4.configure(text='Buscar')
        button4.place(
            anchor="nw",
            height=20,
            relx=0.28,
            rely=0.68,
            width=100,
            x=0,
            y=0)
        button4.configure(command=self.pegarItem)
        self.entry4 = ttk.Entry(frame2)
        self.entry4.place(anchor="nw", relx=0.55, rely=0.64, width=300, x=0, y=0)
        button5 = tk.Button(frame2)
        button5.configure(text='Confimar')
        button5.place(
            anchor="nw",
            height=20,
            relx=0.68,
            rely=0.74,
            width=100,
            x=0,
            y=0)
        button5.configure(command=self.confirmarItem)
        self.entry5 = ttk.Entry(frame2)
        self.entry5.place(anchor="nw", relx=0.8, rely=0.69, width=50, x=0, y=0)
        self.label13 = tk.Label(frame2)
        self.label13.configure(
            background="#E0E0E0",
            font="{ARIAL} 10 {bold}",
            relief="ridge",
            text='Quantidade do item')
        self.label13.place(
            anchor="nw",
            height=20,
            relx=0.6,
            rely=0.69,
            width=150,
            x=0,
            y=0)
        self.label14 = tk.Label(frame2)
        self.label14.configure(
            background="#F5F5F5",
            font="{ARIAL} 10 {bold}",
            text='Item')
        self.label14.place(
            anchor="nw",
            height=20,
            relx=0.63,
            rely=0.61,
            width=150,
            x=0,
            y=0)
        separator7 = ttk.Separator(frame2)
        separator7.configure(orient="vertical")
        separator7.place(
            anchor="nw",
            height=150,
            relx=0.50,
            rely=0.6,
            width=1,
            x=0,
            y=0)
        self.label15 = tk.Label(frame2)
        self.label15.configure(
            background="#E0E0E0",
            font="{ARIAL} 10 {bold}",
            foreground="#ff4246",
            relief="ridge",
            text='Itens já confirmados: ')
        self.label15.place(
            anchor="nw",
            height=20,
            relx=0.38,
            rely=0.83,
            width=200,
            x=0,
            y=0)
        button6 = tk.Button(frame2)
        button6.configure(text='Excluir último item adicionado')
        button6.place(
            anchor="nw",
            height=20,
            relx=0.38,
            rely=0.86,
            width=200,
            x=0,
            y=0)
        button6.configure(command=self.deletarLis)
        self.entry6 = ttk.Entry(frame2)
        self.entry6.place(anchor="nw", relx=0.75, rely=0.605, width=20, x=0, y=0)
        separator2 = ttk.Separator(frame2)
        separator2.configure(orient="horizontal")
        separator2.place(
            anchor="nw",
            relx=0.06,
            rely=0.23,
            width=700,
            x=0,
            y=0)
        self.label12 = tk.Label(frame2)
        self.label12.configure(
            background="#E0E0E0",
            font="{ARIAL} 10 {bold}",
            relief="ridge",
            text='NF')
        self.label12.place(
            anchor="nw",
            height=20,
            relx=0.15,
            rely=0.18,
            width=100,
            x=0,
            y=0)
        self.entry7 = tk.Entry(frame2)
        self.entry7.place(anchor="nw", relx=0.28, rely=0.18, width=100, x=0, y=0)
        self.label16 = tk.Label(frame2)
        self.label16.configure(
            background="#E0E0E0",
            font="{ARIAL} 10 {bold}",
            relief="ridge",
            text='DATA')
        self.label16.place(
            anchor="nw",
            height=20,
            relx=0.5,
            rely=0.18,
            width=100,
            x=0,
            y=0)
        self.entry8 = tk.Entry(frame2)
        self.entry8.place(anchor="nw", relx=0.63, rely=0.18, width=100, x=0, y=0)
        #self.entry6.configure(state='readonly')
        frame2.pack(side="top")

        # Main widget
        self.mainwindow = toplevel1

    def run(self):
        self.mainwindow.mainloop()

    def stop(self):
        self.mainwindow.destroy()

    def pegarCliente(self):
        cliente,x,y,z = Certificado.pegarCliente(self.entry1.get())
        self.label6.configure(text=f"{cliente}",font="{ARIAL} 7 {bold}",relief="ridge")
        self.label7.configure(text=f"{x}",font="{ARIAL} 7 {bold}",relief="ridge")
        self.label8.configure(text=f"{y}",font="{ARIAL} 7 {bold}",relief="ridge")
        return

    def pegarQTDITENS(self):
        qtitens = Certificado.quantosItens(self.entry2.get())
        self.label10.configure(text=f"{qtitens} Itens serão certificados",font="{ARIAL} 10 {bold}",relief="ridge")
        self.entry2.configure(state='readonly')
        return qtitens
    
    def pegarItem(self):
        item = Certificado.pegarItem(self.entry3.get(),self.entry3.get())
        self.entry6.insert(0,"1")
        self.entry6.configure(state='readonly')
        #self.entry3.configure(state='readonly')
        self.entry4.delete(0, 'end')
        self.entry4.insert(0,f'{item[0]}')
        return

    def confirmarItem(self):
        a = int(self.entry6.get())
        b = int(self.entry2.get())
        if a <= b:
            item = Certificado.pegarItem(self.entry3.get(),self.entry3.get())
            item.insert(0,self.entry5.get())
            item.append(self.entry6.get())
            i = len(self.lis)
            self.entry6.configure(state='NORMAL')
            self.entry6.delete(0, 'end')
            self.entry6.insert(0,f'{i+2}')
            self.entry6.configure(state='readonly')
            self.entry3.configure(state='')
            self.entry4.delete(0, 'end')
            self.entry3.delete(0, 'end')
            self.entry5.delete(0, 'end')
            self.lis.append(item)
            self.label15.configure(text=f'Itens já confirmados: {int(self.entry6.get())-1}')
        if a > b:
            HUD2.abre()
        return

    def deletarLis(self):
        self.lis.pop()
        i = len(self.lis)
        i = i+1
        self.entry6.configure(state='NORMAL')
        self.entry6.delete(0,'end')
        self.entry6.insert(0,f'{i}')
        self.entry6.configure(state='readonly')

    def abrirDoc(self):
        Certificado.abrirDoc()

    def gerarLis(self):
        a = Certificado.gerarLis(self.lis)
        w,x,y,z = Certificado.pegarCliente(self.entry1.get())
        Certificado.rodar(w,x,y,z,int(self.entry2.get()),a,self.entry7.get(),self.entry8.get())
        self.mainwindow.destroy()
        HUD3.abre()
        return





def abre():
    app = MUB1()
    app.run()

abre()

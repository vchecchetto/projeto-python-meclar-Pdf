import os
import tkinter as tk
import PyPDF2
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox


class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Mesclar PDF")
        self.root.iconbitmap("pdficon.ico")
        width = 600
        height = 400
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = int((screen_width / 2) - (width / 2))
        y = int((screen_height / 2) - (height / 2))
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.resizable(width=False, height=False)

        def sobre():
            messagebox.showinfo("Sobre", "Software para mesclar PDF."
                                         "\n\nDesenvolvido por Vinicius Checchetto.")

        menubar = Menu(self.root)
        self.root.config(menu=menubar)

        mainmenu = tk.Menu(menubar, tearoff=0)
        mainmenu.add_command(label="Sair", command=self.root.destroy)

        menubar.add_command(label="Sobre", command=sobre)

        self.frame_top = Frame(self.root, width=600, height=80, relief="flat")
        self.frame_top.grid(row=0, column=0, pady=1, padx=0)

        self.frame_down = Frame(self.root, width=600, height=320, relief="flat")
        self.frame_down.grid(row=1, column=0, pady=1, padx=0)

        self.label_titulo = Label(self.frame_top, text="MESCLAR PDF", font="Bahnschrift 26", anchor=NW)
        self.label_titulo.place(x=20, y=30)

        self.arquivos_treeview = ttk.Treeview(self.frame_down, columns=("Arquivo", "Tamanho"), show="headings")
        self.arquivos_treeview.heading("Arquivo", text="Arquivos")
        self.arquivos_treeview.heading("Tamanho", text="Tamanho")
        self.arquivos_treeview.column("Arquivo", width=350)
        self.arquivos_treeview.column("Tamanho", width=100)
        self.arquivos_treeview.place(x=25, y=35, width=500, height=100)

        self.botao_adicionar = Button(self.frame_down, text="Adicionar", font="Calibri 12 bold", width=10, height=1,
                                      relief=RAISED, overrelief=RIDGE, command=self.adicionar_arquivo)
        self.botao_adicionar.place(x=105, y=160)

        self.botao_remover = Button(self.frame_down, text="Remover", font="Calibri 12 bold", width=10, height=1,
                                    relief=RAISED, overrelief=RIDGE, command=self.remover_arquivo)
        self.botao_remover.place(x=225, y=160)

        self.botao_mesclar = Button(self.frame_down, text="Mesclar", font="Calibri 12 bold", width=10, height=1,
                                    relief=RAISED, overrelief=RIDGE, command=self.mesclar_pdf)
        self.botao_mesclar.place(x=345, y=160)

        self.botao_para_cima = Button(self.frame_down, text="↑", font="Calibri 12 bold", width=3, height=1,
                                      relief=RAISED, overrelief=RIDGE, command=self.mover_para_cima)
        self.botao_para_cima.place(x=540, y=40)

        self.botao_para_baixo = Button(self.frame_down, text="↓", font="Calibri 12 bold", width=3, height=1,
                                       relief=RAISED, overrelief=RIDGE, command=self.mover_para_baixo)
        self.botao_para_baixo.place(x=540, y=95)

        self.botao_sair = Button(self.frame_down, text="Sair", font="Calibri 12 bold", width=10, height=1,
                                 relief=RAISED, overrelief=RIDGE, command=self.sair)
        self.botao_sair.place(x=480, y=240)

        self.output_directory = None

    def sair(self):
        self.root.destroy()

    def adicionar_arquivo(self):
        arquivos_pdf = filedialog.askopenfilenames(
            parent=self.frame_down,
            title="Selecionar Arquivo",
            filetypes=(
                ("PDF", "*.pdf"),
                ("Todos os Arquivos", "*.*")
            )
        )
        for arquivo in arquivos_pdf:
            caminho_arquivo = r"{}".format(arquivo)
            tamanho_arquivo = os.path.getsize(caminho_arquivo)
            tamanho_formatado = self.formatar_tamanho(tamanho_arquivo)
            self.arquivos_treeview.insert("", "end", values=(os.path.basename(caminho_arquivo), tamanho_formatado))

            if self.output_directory is None:
                self.output_directory = os.path.dirname(caminho_arquivo)

    def formatar_tamanho(self, tamanho_bytes):
        for unidade in ['B', 'KB', 'MB', 'GB']:
            if tamanho_bytes < 1024.0:
                return "%3.1f %s" % (tamanho_bytes, unidade)
            tamanho_bytes /= 1024.0

    def remover_arquivo(self):
        selected_item = self.arquivos_treeview.selection()
        if selected_item:
            self.arquivos_treeview.delete(selected_item)

    def mover_para_cima(self):
        selected_item = self.arquivos_treeview.selection()
        if selected_item:
            self.arquivos_treeview.move(selected_item, "", self.arquivos_treeview.index(selected_item) - 1)

    def mover_para_baixo(self):
        selected_item = self.arquivos_treeview.selection()
        if selected_item:
            self.arquivos_treeview.move(selected_item, "", self.arquivos_treeview.index(selected_item) + 2)

    def mesclar_pdf(self):
        arquivos_selecionados = self.arquivos_treeview.get_children()
        if len(arquivos_selecionados) < 2:
            messagebox.showwarning("Aviso", "A mesclagem requer pelo menos dois arquivos.")
            return

        mesclar = PyPDF2.PdfMerger()

        for arquivo_id in arquivos_selecionados:
            nome_arquivo = self.arquivos_treeview.item(arquivo_id, "values")[0]
            caminho_arquivo = os.path.join(self.output_directory, nome_arquivo)
            mesclar.append(open(caminho_arquivo, "rb"))

        output_path = filedialog.asksaveasfile(
            initialdir=self.output_directory,
            parent=self.frame_down,
            title="Salvar PDF Mesclado Como...",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")]
        )

        if output_path:
            with open(output_path.name, "wb") as output_file:
                mesclar.write(output_file)
            mesclar.close()
            messagebox.showinfo("Sucesso", "Arquivos PDF mesclados com sucesso!")
        else:
            messagebox.showinfo("Cancelado", "Mesclagem de arquivos PDF cancelada.")

    def janela_principal(self):
        self.root.mainloop()

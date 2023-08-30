#pip install pyinstaller
#pyinstaller --onefile --noconsole doc_tree.py

import os
from reportlab.lib.pagesizes import A4
from reportlab.platypus import PageTemplate, BaseDocTemplate
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph
from tkinter import Tk, filedialog, messagebox, simpledialog, Label, Entry, Button, StringVar, ttk
from reportlab.lib import colors
import datetime

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Árvore de Arquivos")
        
        self.title_var = StringVar()
        self.posto_var = StringVar()
        self.nome_var = StringVar()
        self.source_var = StringVar()
        self.destination_var = StringVar()
        
        self.title_label = Label(root, text="Título do Projeto:")
        self.title_label.pack()
        self.title_entry = Entry(root, textvariable=self.title_var)
        self.title_entry.pack()
        
        # Adicionando os campos de Posto e Nome na mesma linha
        self.posto_label = Label(root, text="Posto:")
        self.posto_label.pack(side="left")
        self.posto_combobox = ttk.Combobox(root, textvariable=self.posto_var, values=[
            "TB", "MB", "BR", "CL", "TC", "MJ", "CP", "1T", "2T", "AP", "SO", "1S", "2S", "3S", "CB", "S1", "S2"
        ])
        self.posto_combobox.pack(side="left")

        self.nome_label = Label(root, text="Nome:")
        self.nome_label.pack(side="left")
        self.nome_entry = Entry(root, textvariable=self.nome_var)
        self.nome_entry.pack(side="left")
        
        self.source_label = Label(root, text="Diretório de Origem:")
        self.source_label.pack()
        self.source_entry = Entry(root, textvariable=self.source_var)
        self.source_entry.pack()
        self.source_button = Button(root, text="Selecionar Diretório de Origem", command=self.select_source)
        self.source_button.pack()
        
        self.destination_label = Label(root, text="Diretório de Destino:")
        self.destination_label.pack()
        self.destination_entry = Entry(root, textvariable=self.destination_var)
        self.destination_entry.pack()
        self.destination_button = Button(root, text="Selecionar Diretório de Destino", command=self.select_destination)
        self.destination_button.pack()
        
        self.generate_button = Button(root, text="Gerar PDF", command=self.generate_pdf)
        self.generate_button.pack()
        
        self.quit_button = Button(root, text="Sair", command=root.destroy)
        self.quit_button.pack()
    
    def select_source(self):
        source_directory = filedialog.askdirectory(title="Selecione o diretório para mapear")
        if source_directory:
            self.source_var.set(source_directory)
    
    def select_destination(self):
        destination_directory = filedialog.askdirectory(title="Selecione o diretório para salvar o PDF")
        if destination_directory:
            self.destination_var.set(destination_directory)
    
    def generate_pdf(self):
        title = self.title_var.get()
        posto = self.posto_var.get()  # Obter valor do campo Posto
        nome = self.nome_var.get()  # Obter valor do campo Nome
        source_directory = self.source_var.get()
        destination_directory = self.destination_var.get()
        
        if title and source_directory and destination_directory:
            sanitized_source_path = source_directory.replace("/", "\\\\")
            sanitized_destination_path = os.path.join(destination_directory.replace("/", "\\\\"), 'ARVORE_DE_ARQUIVOS.pdf')
            
            self.list_files(sanitized_source_path, sanitized_destination_path, title, posto, nome)
            messagebox.showinfo("Informação", "PRONTO!\nArquivo 'ARVORE_DE_ARQUIVOS.pdf' foi gerado.")
        else:
            messagebox.showwarning("Informação", "Preencha todos os campos antes de gerar o PDF.")
    
    def get_excel_style_alphabet_index(self, index):
        alphabet = 'abcdefghijklmnopqrstuvwxyz'
        if index < 26:
            return alphabet[index]
        else:
            first_char = alphabet[(index // 26) - 1]
            second_char = alphabet[index % 26]
            return first_char + second_char

    def list_files(self, startpath, output_path, titulo, posto, nome):
        counters = [0]  # Lista de contadores para cada nível de indentação
        elements = []  # Lista de elementos a serem adicionados ao PDF
        styles = getSampleStyleSheet()
        footer_style = ParagraphStyle('FooterStyle',
                                    parent=styles['BodyText'],
                                    fontName='Helvetica-Oblique',
                                    fontSize=8,
                                    alignment=1,
                                    textColor=colors.grey,
                                    spaceAfter=6)
        folder_style = ParagraphStyle('FolderStyle',
                                    parent=styles['BodyText'],
                                    fontName='Helvetica-Bold',
                                    spaceAfter=6)
        file_style = ParagraphStyle('FileStyle',
                                    parent=styles['BodyText'],
                                    fontName='Helvetica-Oblique',
                                    spaceAfter=0)
        title_style = ParagraphStyle('TitleStyle',
                                    parent=styles['BodyText'],
                                    fontName='Helvetica-Bold',
                                    fontSize=18,  # Definir tamanho da fonte
                                    alignment=1,  # 1 é para alinhamento centralizado
                                    leading=20,
                                    spaceAfter=12)

        titulo_2 = ("Árvore de Arquivos")
        title_style_2 = ParagraphStyle('TitleStyle_2',
                                    parent=styles['BodyText'],
                                    fontName='Helvetica',
                                    fontSize=14,  # Definir tamanho da fonte
                                    alignment=1,  # 1 é para alinhamento centralizado
                                    leading=36,
                                    spaceAfter=12)

        elements.append(Paragraph(f"{titulo}", title_style))
        elements.append(Paragraph(f"{titulo_2}", title_style_2))

        for root, dirs, files in os.walk(startpath):
            level = root.replace(startpath, '').count(os.sep)
            indent = '&nbsp;' * 4 * level
            
            if level + 1 > len(counters):
                counters.extend([0] * (level + 1 - len(counters)))
            counters[level] += 1
            counters[level+1:] = [0] * (len(counters) - level - 1)

            index_str = '.'.join(str(count) for count in counters[:level + 1])

            if level == 0:
                elements.append(Paragraph(f"{os.path.basename(root)}", folder_style))
            else:
                elements.append(Paragraph(f"{indent}{index_str}. {os.path.basename(root)}", folder_style))

            sub_indent = '&nbsp;' * 4 * (level + 1)
            file_counter = 1
            for filename in files:
                if level == 0:
                    file_index = f"{self.get_excel_style_alphabet_index(file_counter-1)}"
                else:
                    file_index = f"{index_str}.{self.get_excel_style_alphabet_index(file_counter-1)}"
                elements.append(Paragraph(f"{sub_indent}{file_index}. {filename}", file_style))
                file_counter += 1

        pdf = SimpleDocTemplate(output_path, pagesize=A4, showBoundary=0)
        footer_text = f"Gerado por {posto} {nome} em {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        footer_paragraph = Paragraph(footer_text, footer_style) 

        elements.append(footer_paragraph)
        pdf.build(elements)
        pass


if __name__ == "__main__":
   root = Tk()
   app = App(root)
   root.mainloop()

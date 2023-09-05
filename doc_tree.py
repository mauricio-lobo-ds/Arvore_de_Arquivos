#Gerar Executável:
#pip install pyinstaller
#pyinstaller doc_tree.spec

import os
from reportlab.lib.pagesizes import A4
from reportlab.platypus import PageTemplate, BaseDocTemplate
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Image
from tkinter import Tk, filedialog, messagebox, simpledialog, Label, Entry, Button, StringVar, ttk, Frame, PhotoImage
from reportlab.lib import colors
import datetime
from PIL import Image, ImageTk
import pystray
import sys

# Classe personalizada para criar um parágrafo com sublinhado
class UnderlinedParagraph(Paragraph):
    def __init__(self, text, style):
        text_with_underline = "<u>{}</u>".format(text)
        super().__init__(text_with_underline, style)

class App:
    def __init__(self, root):
        self.root = root
        self.root.geometry("375x470")
        self.root.title("Gerador de Árvore de Arquivos")
        self.root.resizable(False, False)
        self.background_color = '#c7dded'
        self.root.configure(bg=self.background_color)

        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        dom_path = os.path.join(base_path, 'dom.png')
        icone_path = os.path.join(base_path, 'dom.ico')

        self.root.iconbitmap(icone_path)
        
        self.dom = PhotoImage(file=dom_path)
        self.imagem_dom = Label(image=self.dom)
        self.imagem_dom.pack(pady=8)
        self.imagem_dom.configure(bg=self.background_color)
        
        self.title_var = StringVar()
        self.om_var = StringVar()
        self.posto_var = StringVar()
        self.nome_var = StringVar()
        self.source_var = StringVar()
        self.destination_var = StringVar()
        
        self.title_label = Label(root, text="Título do Projeto:")
        self.title_label.pack()
        self.title_label.configure(bg=self.background_color)
        
        self.title_entry = Entry(root, textvariable=self.title_var, width=56)
        self.title_entry.pack()

        # Criar um Frame para Posto e Nome
        om_frame = Frame(root)
        om_frame.pack(pady=8)

        self.posto_label = Label(om_frame, text="OM Responsável:")
        self.posto_label.pack(side="left")
        self.posto_label.configure(bg=self.background_color)

        self.om_resp = ttk.Combobox(om_frame,textvariable=self.om_var, values=[
            "CEPE", "DTECAMP", "DTINFRA-BE", "DTINFRA-BR", "DTINFRA-CO", "DTINFRA-MN", "DTINFRA-NT", "DTINFRA-RJ", "DTINFRA-SJ"
        ], width=37)
        self.om_resp.pack(side="left")
        # Registre a função de validação
        validate_om = self.root.register(self.validate_posto_om)
        self.om_resp.config(validate="key", validatecommand=(validate_om, "%P"))

        # Criar um Frame para Posto e Nome
        posto_nome_frame = Frame(root)
        posto_nome_frame.pack(pady=8)
        
        self.posto_label = Label(posto_nome_frame, text="Posto:")
        self.posto_label.pack(side="left")
        self.posto_label.configure(bg=self.background_color)

        self.posto_combobox = ttk.Combobox(posto_nome_frame, textvariable=self.posto_var, values=[
            "TB", "MB", "BR", "CL", "TC", "MJ", "CP", "1T", "2T", "AP", "SO", "1S", "2S", "3S", "CB", "S1", "S2"
        ], width=3)
        self.posto_combobox.pack(side="left")
        # Registre a função de validação
        validate_posto = self.root.register(self.validate_posto_om)
        self.posto_combobox.config(validate="key", validatecommand=(validate_posto, "%P"))

        self.nome_label = Label(posto_nome_frame, text="Nome:")
        self.nome_label.pack(side="left")
        self.nome_label.configure(bg=self.background_color)


        self.nome_entry = Entry(posto_nome_frame, textvariable=self.nome_var, width=36)
        self.nome_entry.pack(side="left")
        
        self.source_label = Label(root, text="Diretório de Origem:")
        self.source_label.pack()
        self.source_label.configure(bg=self.background_color)

        self.source_entry = Entry(root, textvariable=self.source_var, width=56)
        self.source_entry.pack()
        

        self.source_button = Button(root, text="Selecionar Diretório de Origem", command=self.select_source, width=27)
        self.source_button.pack(pady=8)
        
        self.generate_button = Button(root, text="Gerar PDF", command=self.generate_pdf, width=27)
        self.generate_button.pack(pady=8)
        
        self.quit_button = Button(root, text="Sair", command=root.destroy, width=27)
        self.quit_button.pack(pady=(5,20))
        
        self.develop_label = Label(root, text="Desenvolvido pela Divisão de Gestão de Engenharia da DIRINFRA", font=("Helvetica",7,"italic"))
        self.develop_label.pack(pady=0)
        self.develop_label.configure(bg=self.background_color)


        self.developer_label = Label(root, text="CP Mauricio e CP Labrego", font=("Helvetica",7,"italic"))
        self.developer_label.pack(pady=0)
        self.developer_label.configure(bg=self.background_color)
        
    
    def validate_posto_om(self, new_value):
        # Verifique se o novo valor inserido está na lista de opções
        if new_value in self.posto_combobox["values"]:
            return True
        else:
            return False
    
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
        om = self.om_var.get()  # Obter valor do campo OM Resp
        posto = self.posto_var.get()  # Obter valor do campo Posto
        nome = self.nome_var.get()  # Obter valor do campo Nome
        source_directory = self.source_var.get()
        destination_directory = self.source_var.get()
        
        try:
            if title and source_directory and destination_directory:
                sanitized_source_path = source_directory.replace("/", "\\\\")
                sanitized_destination_path = os.path.join(destination_directory.replace("/", "\\\\"), 'ARVORE_DE_ARQUIVOS.pdf')
                
                self.list_files(sanitized_source_path, sanitized_destination_path, title, posto, nome, om)
                messagebox.showinfo("Informação", f"PRONTO!\nArquivo 'ARVORE_DE_ARQUIVOS.pdf' foi gerado na pasta:\n{source_directory}")
            else:
                messagebox.showwarning("Informação", "Preencha todos os campos antes de gerar o PDF.")
        except Exception as e:
            print(f"Erro: {str(e)}")
            messagebox.showerror("Erro", "Ocorreu um erro durante a gravação dos dados. Verifique se há algum documento .pdf aberto de árvore de arquivos.")
    
    def get_excel_style_alphabet_index(self, index):
        alphabet = 'abcdefghijklmnopqrstuvwxyz'
        if index < 26:
            return alphabet[index]
        else:
            first_char = alphabet[(index // 26) - 1]
            second_char = alphabet[index % 26]
            return first_char + second_char

    def list_files(self, startpath, output_path, titulo, posto, nome, om):
        counters = [1]  # Lista de contadores para cada nível de indentação
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
        
        MD_style = ParagraphStyle('TitleStyle',
                                    parent=styles['BodyText'],
                                    fontName='Helvetica-Bold',
                                    fontSize=11,  # Definir tamanho da fonte
                                    alignment=1  # 1 é para alinhamento centralizado
                                    )
        
        COMAER_style = ParagraphStyle('TitleStyle',
                                    parent=styles['BodyText'],
                                    fontName='Helvetica',
                                    fontSize=11,  # Definir tamanho da fonte
                                    alignment=1  # 1 é para alinhamento centralizado
                                    )
        
        OM_style = ParagraphStyle('TitleStyle',
                                    parent=styles['BodyText'],
                                    fontName='Helvetica',
                                    fontSize=11,  # Definir tamanho da fonte
                                    alignment=1,  # 1 é para alinhamento centralizado
                                    textDecoration='underline',
                                    spaceAfter=18
                                    )
        
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
                                    leading=24,
                                    spaceAfter=2)

        om_dict = {"DTINFRA-BE":"DESTACAMENTO DE INFRAESTRUTURA DA AERONÁUTICA DE BELÉM",
                   "DTINFRA-BR":"DESTACAMENTO DE INFRAESTRUTURA DA AERONÁUTICA DE BRASÍLIA",
                   "DTINFRA-CO":"DESTACAMENTO DE INFRAESTRUTURA DA AERONÁUTICA DE CANOAS",
                   "DTINFRA-MN":"DESTACAMENTO DE INFRAESTRUTURA DA AERONÁUTICA DE MANAUS",
                   "DTINFRA-NT":"DESTACAMENTO DE INFRAESTRUTURA DA AERONÁUTICA DE NATAL",
                   "DTINFRA-RJ":"DESTACAMENTO DE INFRAESTRUTURA DA AERONÁUTICA DO RIO DE JANEIRO",
                   "DTINFRA-SJ":"DESTACAMENTO DE INFRAESTRUTURA DA AERONÁUTICA DE SÃO JOSÉ DOS CAMPOS",
                   "DTECAMP":"DESTACAMENTO DE ENGENHARIA DE CAMPANHA",
                   "CEPE": "CENTRO DE ESTUDOS E PROJETOS DE ENGENHARIA DA AERONÁUTICA"}
        
        elements.append(Paragraph("MINISTÉRIO DA DEFESA", MD_style))
        elements.append(Paragraph("COMANDO DA AERONÁUTICA", COMAER_style))
        om_sublinhada = UnderlinedParagraph(f"{om_dict[om]}",OM_style)
        elements.append(om_sublinhada)
        elements.append(Paragraph(f"{titulo}", title_style))
        elements.append(Paragraph(f"{titulo_2}", title_style_2))

        footer_text = f"Gerado por {posto} {nome} em {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        footer_paragraph = Paragraph(footer_text, footer_style) 

        elements.append(footer_paragraph)

        for root, dirs, files in os.walk(startpath):
            level = root.replace(startpath, '').count(os.sep)
            indent = '&nbsp;' * 4 * level
            
            if level == 0:
                elements.append(Paragraph(f"{os.path.basename(root)}", folder_style))
            else:
                if level + 1 > len(counters):
                    counters.extend([0] * (level + 1 - len(counters)))
                counters[level] += 1
                counters[level+1:] = [0] * (len(counters) - level - 1)

                index_str = '.'.join(str(count) for count in counters[1:level + 1])

                elements.append(Paragraph(f"{indent}[{index_str}]   {os.path.basename(root)}", folder_style))

            sub_indent = '&nbsp;' * 4 * (level + 1)
            file_counter = 1
            flag_treefile = 0
            for filename in files:
                if level == 0:
                    file_index = f"{self.get_excel_style_alphabet_index(file_counter-1)}"
                else:
                    file_index = f"{index_str}.{self.get_excel_style_alphabet_index(file_counter-1)}"
                elements.append(Paragraph(f"{sub_indent}[{file_index}]   {filename}", file_style))
                if filename == "ARVORE_DE_ARQUIVOS.pdf":
                    flag_treefile = 1
                file_counter += 1
               
            if level == 0:
                file_index = f"{self.get_excel_style_alphabet_index(file_counter-1)}"
                if flag_treefile == 0:
                    elements.append(Paragraph(f"{sub_indent}[{file_index}]   ARVORE_DE_ARQUIVOS.pdf", file_style))

        pdf = SimpleDocTemplate(output_path, pagesize=A4, showBoundary=0, leftMargin=50, rightMargin=30, topMargin=30, bottomMargin=30)
        
        pdf.build(elements)
        pass

if __name__ == "__main__":
   root = Tk()
   app = App(root)
   root.mainloop()

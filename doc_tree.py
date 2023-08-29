#pip install pyinstaller
#pyinstaller --onefile --noconsole doc_tree.py

import os
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph
from tkinter import Tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import simpledialog


def ask_for_directory(prompt):
    root = Tk()  # Cria uma janela Tkinter. Isso é necessário para abrir o diálogo.
    root.withdraw()  # Esconde a janela Tkinter, para que apenas o diálogo seja mostrado.
    messagebox.showinfo("Informação", prompt)  # Exibe a caixa de diálogo de informação
    folder_selected = filedialog.askdirectory(title=prompt)  # Abre o diálogo de seleção de pasta.
    return folder_selected

def ask_for_title(prompt):
    root = Tk()
    root.withdraw()
    return simpledialog.askstring(" ", prompt)

def get_excel_style_alphabet_index(index):
    alphabet = 'abcdefghijklmnopqrstuvwxyz'
    if index < 26:
        return alphabet[index]
    else:
        first_char = alphabet[(index // 26) - 1]
        second_char = alphabet[index % 26]
        return first_char + second_char

def list_files(startpath, output_path):
    counters = [0]  # Lista de contadores para cada nível de indentação
    elements = []  # Lista de elementos a serem adicionados ao PDF
    styles = getSampleStyleSheet()
    folder_style = ParagraphStyle('FolderStyle',
                                  parent=styles['BodyText'],
                                  fontName='Helvetica-Bold',
                                  spaceAfter=6)
    file_style = ParagraphStyle('FileStyle',
                                parent=styles['BodyText'],
                                fontName='Helvetica-Oblique',
                                spaceAfter=0)
    
    titulo = ask_for_title("Insira o nome do Projeto")
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

        elements.append(Paragraph(f"{indent}{index_str}. {os.path.basename(root)}", folder_style))

        sub_indent = '&nbsp;' * 4 * (level + 1)
        file_counter = 1
        for filename in files:
            file_index = f"{index_str}.{get_excel_style_alphabet_index(file_counter-1)}"
            elements.append(Paragraph(f"{sub_indent}{file_index}. {filename}", file_style))
            file_counter += 1

    pdf = SimpleDocTemplate(output_path, pagesize=A4)
    pdf.build(elements)

if __name__ == "__main__":
    source_directory = ask_for_directory("Selecione o diretório para mapear")

    if source_directory:
        sanitized_source_path = source_directory.replace("/", "\\\\")
        print(f"Gerando a estrutura de arquivos para {sanitized_source_path}...")
        
        try:
            destination_directory = ask_for_directory("Selecione o diretório para salvar o PDF")
            if destination_directory:
                sanitized_destination_path = os.path.join(destination_directory.replace("/", "\\\\"), 'ARVORE_DE_ARQUIVOS.pdf')
                print(f"Salvando o PDF em {sanitized_destination_path}...")
                
                list_files(sanitized_source_path, sanitized_destination_path)  # Certifique-se de adaptar sua função list_files para aceitar o destino do PDF como um argumento.
                print("Arquivo 'ARVORE_DE_ARQUIVOS.pdf' foi gerado.")
                messagebox.showinfo("Informação", "Arquivo 'ARVORE_DE_ARQUIVOS.pdf' foi gerado.")
            else:
                print("Nenhum diretório de destino selecionado.")
                messagebox.showwarning("Informação","Nenhum diretório de destino selecionado.")
        except:
            messagebox.showerror("Informação","Nao foi possível salvar o documento. Verifique se há algum documento .pdf aberto de arvore de arquivos.")
    else:
        print("Nenhum diretório de origem selecionado.")
        messagebox.showwarning("Informação","Nenhum diretório de origem selecionado.")
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import pandas as pd
import webbrowser

class BoardGamersPortugalApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Visualizador de Dados Excel Local")
        self.root.geometry("800x500")

        # Nome do arquivo Excel local
        self.excel_file = "Boardgamers em Portugal_BD.xlsx"

        # Carregar os dados do Excel
        self.sheets_dict = self.load_data()

        # Selecionar sheet
        self.sheet_var = tk.StringVar()
        self.sheet_dropdown = ttk.Combobox(root, textvariable=self.sheet_var, values=list(self.sheets_dict.keys()))
        self.sheet_dropdown.set("Escolha uma sheet")
        self.sheet_dropdown.bind("<<ComboboxSelected>>", self.update_table)
        self.sheet_dropdown.pack(pady=5)

        # Campos para filtro
        self.filter_label = tk.Label(root, text="Filtrar por:")
        self.filter_label.pack(pady=5)

        # Opções de coluna para filtro (inicialmente vazias, atualizadas ao selecionar sheet)
        self.column_var = tk.StringVar()
        self.column_dropdown = ttk.Combobox(root, textvariable=self.column_var)
        self.column_dropdown.pack(pady=5)

        # Campo de entrada do valor de filtro
        self.filter_entry = tk.Entry(root)
        self.filter_entry.pack(pady=5)

        # Botão para aplicar o filtro
        self.filter_button = tk.Button(root, text="Aplicar Filtro", command=self.apply_filter)
        self.filter_button.pack(pady=5)

        # Frame para a Tabela
        self.table_frame = tk.Frame(root)
        self.table_frame.pack(fill="both", expand=True)

        # Exibe a tabela inicial sem filtros (caso uma sheet tenha sido selecionada)
        self.display_table(pd.DataFrame())

    def load_data(self):
        try:
            # Carrega todas as sheets do Excel como dicionário de DataFrames
            xls = pd.ExcelFile(self.excel_file)
            sheets_dict = {sheet: xls.parse(sheet).fillna('') for sheet in xls.sheet_names}
            return sheets_dict
        except FileNotFoundError:
            messagebox.showerror("Erro", f"Arquivo '{self.excel_file}' não encontrado.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar o arquivo:\n{e}")
        return {}

    def update_table(self, event=None):
        # Obtém a sheet selecionada e atualiza as opções de coluna
        selected_sheet = self.sheet_var.get()
        if selected_sheet in self.sheets_dict:
            self.df_original = self.sheets_dict[selected_sheet]
            self.column_dropdown['values'] = list(self.df_original.columns)
            self.column_var.set("Escolha uma coluna")
            self.display_table(self.df_original)

    def apply_filter(self):
        # Obtém a coluna selecionada e o valor de filtro
        selected_column = self.column_var.get()
        filter_value = self.filter_entry.get()

        # Verifica se a coluna e o valor do filtro são válidos
        if selected_column not in self.df_original.columns:
            messagebox.showwarning("Aviso", "Por favor, escolha uma coluna válida para filtrar.")
            return

        # Aplica o filtro ao DataFrame
        filtered_df = self.df_original[self.df_original[selected_column].astype(str).str.contains(filter_value, na=False, case=False)]

        # Exibe a tabela filtrada
        self.display_table(filtered_df)

    def display_table(self, df):
        # Limpa a tabela antiga, se houver
        for widget in self.table_frame.winfo_children():
            widget.destroy()

        if df.empty:
            return

        # Cria o widget Text para exibir os dados
        text_widget = tk.Text(self.table_frame, wrap="none")
        text_widget.pack(fill="both", expand=True)

        # Adiciona cabeçalhos das colunas
        headers = " | ".join(df.columns) + "\n"
        text_widget.insert("end", headers)
        text_widget.insert("end", "-" * len(headers) + "\n")

        # Insere os dados linha por linha, transformando URLs em links
        for _, row in df.iterrows():
            row_text = []
            for value in row:
                if isinstance(value, str) and (value.startswith("http") or "maps.google.com" in value):
                    # Formata o link e adiciona como clicável
                    start = text_widget.index("end")
                    text_widget.insert("end", value + " ")
                    text_widget.tag_add(value, start, text_widget.index("end-1c"))
                    text_widget.tag_bind(value, "<Button-1>", lambda e, v=value: self.open_link(v))
                    text_widget.tag_config(value, foreground="blue", underline=True)
                else:
                    row_text.append(str(value))
            text_widget.insert("end", " | ".join(row_text) + "\n")

        # Adiciona uma barra de rolagem
        scrollbar = ttk.Scrollbar(self.table_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

    def open_link(self, url):
        webbrowser.open(url)

# Inicializa a aplicação
if __name__ == "__main__":
    root = tk.Tk()
    app = BoardGamersPortugalApp(root)
    root.mainloop()

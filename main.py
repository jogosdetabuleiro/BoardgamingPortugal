import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import pandas as pd

class BoardGamersPortugalApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Visualizador de Dados Excel Local")
        self.root.geometry("800x500")

        # Nome do arquivo Excel local
        self.excel_file = "Boardgamers em Portugal_BD.xlsx"
        
        # Carregar as sheets (abas) do Excel
        self.sheets = self.get_sheets()
        
        # Dropdown para selecionar a aba do Excel
        self.sheet_var = tk.StringVar()
        self.sheet_var.set("Escolha uma aba")
        self.sheet_dropdown = ttk.Combobox(root, textvariable=self.sheet_var, values=self.sheets)
        self.sheet_dropdown.pack(pady=5)
        
        # Botão para carregar e exibir os dados da aba selecionada
        self.load_button = tk.Button(root, text="Carregar Dados da Aba Selecionada", command=self.load_data)
        self.load_button.pack(pady=10)

        # Campos para filtro
        self.filter_label = tk.Label(root, text="Filtrar por:")
        self.filter_label.pack(pady=5)

        # Opções de coluna para filtro (atualizado após carregamento dos dados)
        self.column_var = tk.StringVar()
        self.column_var.set("Escolha uma coluna")
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

    def get_sheets(self):
        """Obtém as sheets (abas) do arquivo Excel."""
        try:
            xls = pd.ExcelFile(self.excel_file)
            return xls.sheet_names
        except FileNotFoundError:
            messagebox.showerror("Erro", f"Arquivo '{self.excel_file}' não encontrado.")
            return []
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar as sheets do arquivo:\n{e}")
            return []

    def load_data(self):
        """Carrega os dados da aba selecionada e exibe na interface."""
        selected_sheet = self.sheet_var.get()
        if selected_sheet not in self.sheets:
            messagebox.showwarning("Aviso", "Por favor, selecione uma aba válida.")
            return

        try:
            # Carrega o arquivo Excel com a aba selecionada, removendo linhas vazias e substituindo NaN por string vazia
            df = pd.read_excel(self.excel_file, sheet_name=selected_sheet, engine='openpyxl').dropna(how="all").fillna("")
            self.df_original = df  # Guarda o DataFrame original

            # Atualiza as opções de coluna para o filtro
            self.column_dropdown["values"] = list(df.columns)

            # Exibe a tabela inicial sem filtros
            self.display_table(df)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar a aba:\n{e}")

    def apply_filter(self):
        """Aplica um filtro ao DataFrame carregado e exibe a tabela filtrada."""
        selected_column = self.column_var.get()
        filter_value = self.filter_entry.get()

        if selected_column not in self.df_original.columns:
            messagebox.showwarning("Aviso", "Por favor, escolha uma coluna válida para filtrar.")
            return

        # Aplica o filtro ao DataFrame
        filtered_df = self.df_original[self.df_original[selected_column].astype(str).str.contains(filter_value, na=False, case=False)]

        # Exibe a tabela filtrada
        self.display_table(filtered_df)

    def display_table(self, df):
        """Exibe o DataFrame em um Treeview."""
        # Limpa a tabela antiga, se houver
        for widget in self.table_frame.winfo_children():
            widget.destroy()

        # Cria o widget Treeview para exibir os dados
        tree = ttk.Treeview(self.table_frame)
        tree.pack(fill="both", expand=True)

        # Configura as colunas no Treeview
        tree["columns"] = list(df.columns)
        tree["show"] = "headings"  # Remove a primeira coluna vazia

        # Define os cabeçalhos das colunas
        for column in df.columns:
            tree.heading(column, text=column)
            tree.column(column, anchor="center")

        # Insere os dados linha por linha
        for _, row in df.iterrows():
            tree.insert("", "end", values=list(row))

        # Adiciona uma barra de rolagem
        scrollbar = ttk.Scrollbar(self.table_frame, orient="vertical", command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

# Inicializa a aplicação
if __name__ == "__main__":
    root = tk.Tk()
    app = BoardGamersPortugalApp(root)
    root.mainloop()

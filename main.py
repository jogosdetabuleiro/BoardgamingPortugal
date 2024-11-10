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

        # Carregar os dados do Excel
        self.df_original = self.load_data()

        # Campos para filtro
        self.filter_label = tk.Label(root, text="Filtrar por:")
        self.filter_label.pack(pady=5)

        # Opções de coluna para filtro
        self.column_var = tk.StringVar()
        self.column_var.set("Escolha uma coluna")
        self.column_dropdown = ttk.Combobox(root, textvariable=self.column_var, values=list(self.df_original.columns))
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

        # Exibe a tabela inicial sem filtros
        self.display_table(self.df_original)

    def load_data(self):
        try:
            # Carrega o arquivo Excel em um DataFrame, removendo linhas totalmente vazias e substituindo NaN por string vazia
            df = pd.read_excel(self.excel_file, engine='openpyxl').dropna(how="all").fillna("")
            return df
        except FileNotFoundError:
            messagebox.showerror("Erro", f"Arquivo '{self.excel_file}' não encontrado.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar o arquivo:\n{e}")
        return pd.DataFrame()  # Retorna um DataFrame vazio em caso de erro

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

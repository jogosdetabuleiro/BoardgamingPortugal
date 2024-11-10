import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import pandas as pd

class BoardGamersPortugalApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Visualizador de Dados Excel Local")
        self.root.geometry("800x400")

        # Nome do arquivo Excel local
        self.excel_file = "Boardgamers em Portugal_BD.xlsx"

        # Botão para carregar e exibir os dados
        self.load_button = tk.Button(root, text="Carregar Dados do Excel Local", command=self.load_data)
        self.load_button.pack(pady=10)

        # Frame para a Tabela
        self.table_frame = tk.Frame(root)
        self.table_frame.pack(fill="both", expand=True)

    def load_data(self):
        try:
            # Carrega o arquivo Excel em um DataFrame
            df = pd.read_excel(self.excel_file, engine='openpyxl')

            # Limpa a tabela antiga, se houver
            for widget in self.table_frame.winfo_children():
                widget.destroy()

            # Exibe os dados do DataFrame em uma tabela
            self.display_table(df)
        except FileNotFoundError:
            messagebox.showerror("Erro", f"Arquivo '{self.excel_file}' não encontrado.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar o arquivo:\n{e}")

    def display_table(self, df):
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

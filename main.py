import tkinter as tk
from tkinter import messagebox
import pandas as pd
import requests
from io import BytesIO

class OnlineExcelViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Visualizador de Dados Excel Online")
        self.root.geometry("600x400")

        # URL do arquivo Excel
        self.excel_url = "https://tinyurl.com/DocBGs"

        # Botão para carregar e exibir os dados
        self.load_button = tk.Button(root, text="Boardgaming em Portugal", command=self.load_data)
        self.load_button.pack(pady=10)

        # Frame para a Tabela
        self.table_frame = tk.Frame(root)
        self.table_frame.pack(fill="both", expand=True)

    def load_data(self):
        try:
            # Faz o download do arquivo Excel da URL
            response = requests.get(self.excel_url)
            response.raise_for_status()  # Verifica se houve erro na requisição

            # Carrega o arquivo Excel em um DataFrame
            excel_data = BytesIO(response.content)
            df = pd.read_excel(excel_data, engine='openpyxl')

            # Limpa a tabela antiga, se houver
            for widget in self.table_frame.winfo_children():
                widget.destroy()

            # Exibe os dados do DataFrame em uma tabela
            self.display_table(df)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar o arquivo:\n{e}")

    def display_table(self, df):
        # Cria uma barra de rolagem para a tabela
        scrollbar = tk.Scrollbar(self.table_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Cria uma lista para exibir os dados
        listbox = tk.Listbox(self.table_frame, yscrollcommand=scrollbar.set)
        listbox.pack(side=tk.LEFT, fill="both", expand=True)

        # Define os cabeçalhos da tabela
        headers = df.columns.tolist()
        listbox.insert(tk.END, " | ".join(headers))

        # Insere os dados linha por linha
        for _, row in df.iterrows():
            row_data = " | ".join(map(str, row.values))
            listbox.insert(tk.END, row_data)

        scrollbar.config(command=listbox.yview)

# Inicializa a aplicação
if __name__ == "__main__":
    root = tk.Tk()
    app = BoardgamingPortugalApp(root)
    root.mainloop()

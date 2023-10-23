import tkinter as tk
from tkinter import ttk
import pandas as pd
from pandastable import Table, TableModel

class TabelaDinamica:
    def __init__(self, root):
        self.root = root
        self.root.title("Tabela Dinâmica - SERGIO HENRIQUE REIS SA")

        frame = ttk.Frame(root)
        frame.pack(padx=10, pady=10)

        # Dados de exemplo (substitua isso pelos seus dados reais)
        data = {
            "Nome": ["João", "Maria", "Carlos", "Ana"],
            "Idade": [25, 32, 45, 28],
            "Cidade": ["São Paulo", "Rio de Janeiro", "Curitiba", "Belo Horizonte"]
        }

        df = pd.DataFrame(data)
        self.table = Table(frame, dataframe=df, showtoolbar=True, showstatusbar=True)
        self.table.model = TableModel(df)

        # Personalizando a tabela
        self.table.configure(bg="lightgray")  # Cor de fundo
        self.table.autoResizeColumns()
        self.table.show()

        frame.pack(side="top", fill="both", expand=True)
        frame.config(width=600, height=400)  # Tamanho da tabela

if __name__ == "__main__":
    root = tk.Tk()
    app = TabelaDinamica(root)
    root.mainloop()

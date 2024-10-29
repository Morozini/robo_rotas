import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import filedialog
from convercao2 import executar


class convercaoVersaodois:

    def __init__(self):
        self.window = tk.Tk()
        self.bg_color = "#BFDDF3"
        self.tela()
        self.tela_principal()

    def tela(self):

        self.window.title("conversion")
        self.window.resizable(False,False)

        largura_tela = self.window.winfo_screenwidth()
        altura_tela = self.window.winfo_screenheight()

        # Definir as dimensões da janela
        largura_janela = 420
        altura_janela = 450

        # Calcular a posição x e y para centralizar a janela
        x = (largura_tela // 2) - (largura_janela // 2)
        y = (altura_tela // 2) - (altura_janela // 2)

        # Definir a geometria da janela
        self.window.geometry(f"{largura_janela}x{altura_janela}+{x}+{y}")

        self.window.configure(background=self.bg_color)  # Define a cor de fundo da janela

    def selecionar_arquivo(self):
        # Abre a janela de diálogo para selecionar um arquivo
        self.caminho_arquivo = filedialog.askopenfilename(
            title="Selecione um arquivo",
            filetypes=(("Arquivos de Texto", "*.kml"), ("Todos os Arquivos", "*.*"))
        )
        if self.caminho_arquivo:  # Verifica se um arquivo foi selecionado
            print(f"Arquivo selecionado: {self.caminho_arquivo}")
            self.caminho_arquivo_var.set(self.caminho_arquivo)

        else:
            print("Nenhum arquivo selecionado.")

    def tela_principal(self):

        self.caminho_arquivo_var = tk.StringVar()

        # Definir cores
        
        text_color = "#000000" 

        list_uf = ['AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MT', 'MS', 'MG', 'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN', 'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO']

        estilo = ttk.Style()
        estilo.configure("My.TFrame", background="#FFFFFF")

        frame = ttk.Frame(self.window, style="My.TFrame")
        frame.pack(pady=20, ipadx=70)

        # Frame 1
        user_info_frame = ttk.Frame(frame, style="My.TFrame")
        user_info_frame.grid(row=0, column=0, padx=20, pady=30)

        title = ttk.Label(user_info_frame, text="Selecione sua UF", foreground=text_color, background="#FFFFFF")
        title.grid(row=0, column=1, padx=10, pady=5)

        self.title_combobox = ttk.Combobox(user_info_frame, values=list_uf)
        self.title_combobox.grid(row=1, column=1, padx=10, pady=5)

        # Frame 2

        kml_name_frame = ttk.Frame(frame, style="My.TFrame")
        kml_name_frame.grid(row=1, column=0, padx=20, pady=30)

        kml_name_arquivo = ttk.Label(kml_name_frame, text="Digite o caminho do KML", background="#FFFFFF")
        kml_name_arquivo.grid(row=1, column=0, sticky="new", padx=5, pady=5)

        self.botao = ttk.Button(kml_name_frame, text="Selecionar Arquivo", command=self.selecionar_arquivo)
        self.botao.grid(row=2, column=0, padx=5, pady=5)


        # Frame 3

        kml_save_frame = ttk.Frame(frame, style="My.TFrame")
        kml_save_frame.grid(row=2, column=0, padx=20, pady=30)

        kml_save_arquivo = ttk.Label(kml_save_frame, text="Digite o caminho para salvar o KML", background="#FFFFFF")
        kml_save_arquivo.grid(row=1, column=0, sticky="new", padx=5, pady=5)

        self.kml_save_input = ttk.Entry(kml_save_frame)
        self.kml_save_input.grid(row=2, column=0, padx=5, pady=5)


        self.botao_convercao = ttk.Button(frame, text="Iniciar", command=self.execute)
        self.botao_convercao.grid(row=3, column=0, padx=5, pady=5)

        # Criar uma label de texto na parte inferior
        bottom_label = ttk.Label(frame, text="Desenvolvido por [Gabriel Morozini]", font=('Arial', 8), foreground=text_color,background="#FFFFFF")
        bottom_label.grid(row=10, column=0, pady=10)

        # Centralizar tudo no grid principal
        frame.grid_columnconfigure(0, weight=1)

    def execute(self):
        arquivo =self.caminho_arquivo
        caminho_arquivo=self.kml_save_input.get()
        nome_saida = 'MAPA_ROTA'
        parse_nome= self.title_combobox.get()

        values = [arquivo, caminho_arquivo, nome_saida, parse_nome]
        executar(values)

        messagebox.showinfo("Sucesso", "Foi finalizado com sucesso")


if __name__  == "__main__":
    app = convercaoVersaodois()
    app.window.mainloop()
    
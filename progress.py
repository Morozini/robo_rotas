import customtkinter as ctk
import time

# Cria a janela principal
root = ctk.CTk()

# Define o título e o tamanho da janela
root.title("Barra de Progresso")
root.geometry("300x150")

# Função para atualizar a barra de progresso
def atualizar_barra():
    progress_bar.set(0)  # Reseta a barra
    for i in range(101):  # Atualiza a barra de 0 a 100%
        progress_bar.set(i / 100)  # Define a porcentagem da barra
        root.update_idletasks()  # Atualiza a interface
        time.sleep(0.05)  # Pausa para efeito de progresso

# Cria a barra de progresso
progress_bar = ctk.CTkProgressBar(root)
progress_bar.pack(pady=20, padx=20)  # Adiciona a barra na interface

# Botão para iniciar a barra de progresso
start_button = ctk.CTkButton(root, text="Iniciar", command=atualizar_barra)
start_button.pack(pady=20)

# Executa a interface
root.mainloop()

import tkinter as tk
from tkinter import messagebox
import threading
import main

def iniciar():
    valor = entrada.get()

    if not valor.isdigit():
        messagebox.showerror("Erro", "Digite apenas números")
        return
    
    num = int(valor)

    threading.Thread(target=main.main, args=(num,), daemon=True).start()

def pause():
    messagebox.showinfo(
        "código pausado"
    )
    # Aqui você chama sua função principal
    # executar_automacao(registros)

# Janela principal
janela = tk.Tk()
janela.title("Automação SIGAMA")
janela.geometry("400x200")
janela.resizable(False, False)

# Label
tk.Label(
    janela,
    text="Quantidade de registros:",
    font=("Arial", 11)
).pack(pady=10)

# Entrada
entrada = tk.Entry(janela, font=("Arial", 11), justify="center")
entrada.pack()

# Botão
tk.Button(
    janela,
    text="Iniciar",
    font=("Arial", 11),
    width=15,
    command=iniciar
).pack(pady=20)

#tk.Button(janela, text="Pausar", font=("Arial", 11), width=15, command=pause).pack(pady=20)

janela.mainloop()

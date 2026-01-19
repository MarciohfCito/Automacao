import tkinter as tk
from tkinter import messagebox
import threading
import main
from control import pause_event
import tkinter as tk
from tkinter import messagebox
import threading
import main
from control import pause_event
from hotkey import iniciar_listener

# ===============================
# FUNÇÕES
# ===============================

def iniciar():
    valor = entrada.get().strip()

    if not valor.isdigit():
        messagebox.showerror("Erro", "Digite apenas números")
        return

    num = int(valor)

    # 🔹 troca de tela
    frame_inicio.pack_forget()
    frame_execucao.pack(fill="both", expand=True)

    janela.title("Executando automação...")

    # 🔹 inicia a main em thread
    threading.Thread(
        target=executar_main,
        args=(num,),
        daemon=True
    ).start()


def executar_main(num):
    try:
        main.main(num)  # <-- SUA AUTOMAÇÃO
    except Exception as e:
        janela.after(0, lambda: messagebox.showerror("Erro", str(e)))
    finally:
        # 🔹 quando terminar, volta para interface
        janela.after(0, finalizar_execucao)


def finalizar_execucao():
    frame_execucao.pack_forget()
    frame_inicio.pack(fill="both", expand=True)

    janela.title("Automação SIGAMA")
    messagebox.showinfo("Concluído", "Automação finalizada com sucesso!")

def pausar_continuar():
    if pause_event.is_set():
        pause_event.clear()
        btn_pausar.config(text="Continuar")
        status_label.config(text="Automação pausada")
    else:
        pause_event.set()
        btn_pausar.config(text="Pausar")
        status_label.config(text="Executando automação...")

# ===============================
# JANELA PRINCIPAL
# ===============================

janela = tk.Tk()
janela.title("Automação SIGAMA")
janela.geometry("400x220")
janela.resizable(False, False)

# ⬇️ inicia listener de atalho global
threading.Thread(
    target=iniciar_listener,
    daemon=True
).start()

# ===============================
# FRAME INICIAL
# ===============================

frame_inicio = tk.Frame(janela)

tk.Label(
    frame_inicio,
    text="Quantidade de registros:",
    font=("Arial", 11)
).pack(pady=15)

entrada = tk.Entry(
    frame_inicio,
    font=("Arial", 11),
    justify="center",
    width=20
)
entrada.pack()

btn_iniciar = tk.Button(
    frame_inicio,
    text="Iniciar",
    font=("Arial", 11),
    width=15,
    command=iniciar
)
btn_iniciar.pack(pady=25)

frame_inicio.pack(fill="both", expand=True)


# ===============================
# FRAME EXECUÇÃO
# ===============================

frame_execucao = tk.Frame(janela)

tk.Label(
    frame_execucao,
    text="Executando automação...",
    font=("Arial", 12, "bold")
).pack(pady=40)

tk.Label(
    frame_execucao,
    text="Por favor, não utilize o computador",
    font=("Arial", 10)
).pack()

status_label = tk.Label(
    frame_execucao,
    text="Executando automação...",
    font=("Arial", 11, "bold")
)
status_label.pack(pady=20)

btn_pausar = tk.Button(
    frame_execucao,
    text="Pausar",
    width=15,
    command=pausar_continuar
)
btn_pausar.pack(pady=10)

frame_execucao.pack_forget()

# ===============================
# LOOP PRINCIPAL
# ===============================

janela.mainloop()


# ===============================
# FUNÇÕES
# ===============================

def iniciar():
    valor = entrada.get().strip()

    if not valor.isdigit():
        messagebox.showerror("Erro", "Digite apenas números")
        return

    num = int(valor)

    # 🔹 troca de tela
    frame_inicio.pack_forget()
    frame_execucao.pack(fill="both", expand=True)

    janela.title("Executando automação...")

    # 🔹 inicia a main em thread
    threading.Thread(
        target=executar_main,
        args=(num,),
        daemon=True
    ).start()


def executar_main(num):
    try:
        main.main(num)  # <-- SUA AUTOMAÇÃO
    except Exception as e:
        janela.after(0, lambda: messagebox.showerror("Erro", str(e)))
    finally:
        # 🔹 quando terminar, volta para interface
        janela.after(0, finalizar_execucao)


def finalizar_execucao():
    frame_execucao.pack_forget()
    frame_inicio.pack(fill="both", expand=True)

    janela.title("Automação SIGAMA")
    messagebox.showinfo("Concluído", "Automação finalizada com sucesso!")

def pausar_continuar():
    if pause_event.is_set():
        pause_event.clear()
        btn_pausar.config(text="Continuar")
        status_label.config(text="Automação pausada")
    else:
        pause_event.set()
        btn_pausar.config(text="Pausar")
        status_label.config(text="Executando automação...")

# ===============================
# JANELA PRINCIPAL
# ===============================

janela = tk.Tk()
janela.title("Automação SIGAMA")
janela.geometry("400x220")
janela.resizable(False, False)


# ===============================
# FRAME INICIAL
# ===============================

frame_inicio = tk.Frame(janela)

tk.Label(
    frame_inicio,
    text="Quantidade de registros:",
    font=("Arial", 11)
).pack(pady=15)

entrada = tk.Entry(
    frame_inicio,
    font=("Arial", 11),
    justify="center",
    width=20
)
entrada.pack()

btn_iniciar = tk.Button(
    frame_inicio,
    text="Iniciar",
    font=("Arial", 11),
    width=15,
    command=iniciar
)
btn_iniciar.pack(pady=25)

frame_inicio.pack(fill="both", expand=True)


# ===============================
# FRAME EXECUÇÃO
# ===============================

frame_execucao = tk.Frame(janela)

tk.Label(
    frame_execucao,
    text="Executando automação...",
    font=("Arial", 12, "bold")
).pack(pady=40)

tk.Label(
    frame_execucao,
    text="Por favor, não utilize o computador",
    font=("Arial", 10)
).pack()

status_label = tk.Label(
    frame_execucao,
    text="Executando automação...",
    font=("Arial", 11, "bold")
)
status_label.pack(pady=20)

btn_pausar = tk.Button(
    frame_execucao,
    text="Pausar",
    width=15,
    command=pausar_continuar
)
btn_pausar.pack(pady=10)

frame_execucao.pack_forget()

# ===============================
# LOOP PRINCIPAL
# ===============================

janela.mainloop()

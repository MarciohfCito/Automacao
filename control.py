import threading

pause_event = threading.Event()
stop_event = threading.Event()

# começa liberado
pause_event.set()

def check_pause():
    pause_event.wait()
    if stop_event.is_set():
        raise SystemExit("Execução interrompida")
from pynput import keyboard
from control import pause_event, stop_event

ctrl = False
alt = False

def toggle_pause():
    if pause_event.is_set():
        pause_event.clear()
        print("⏸️ Automação pausada")
    else:
        pause_event.set()
        print("▶️ Automação retomada")

def stop():
    stop_event.set()
    pause_event.set()
    print("🛑 Automação encerrada")

def on_press(key):
    global ctrl, alt

    if key in (keyboard.Key.ctrl_l, keyboard.Key.ctrl_r):
        ctrl = True

    if key in (keyboard.Key.alt_l, keyboard.Key.alt_r):
        alt = True

    try:
        if ctrl and alt and key.char == 'p':
            toggle_pause()

        if ctrl and alt and key.char == 's':
            stop()
    except AttributeError:
        pass

def on_release(key):
    global ctrl, alt

    if key in (keyboard.Key.ctrl_l, keyboard.Key.ctrl_r):
        ctrl = False

    if key in (keyboard.Key.alt_l, keyboard.Key.alt_r):
        alt = False

def iniciar_listener():
    with keyboard.Listener(
        on_press=on_press,
        on_release=on_release
    ) as listener:
        listener.join()

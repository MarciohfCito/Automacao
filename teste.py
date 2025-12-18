import pyautogui
import pyperclip
import time

#inserir o numero de registros
num = int(input("Digite o número de registros: "))

def minimizar():
    pyautogui.moveTo(pyautogui.moveTo(x=1806, y=7))
    pyautogui.click()
    time.sleep(0.5)

#minimizar vscode
minimizar()
time.sleep(0.5)

#POSIÇÕES
posicao_i_nome_S = [77, 336]
posicao_a_nome_S = posicao_i_nome_S

posicao_i_nome_E = [252, 309]
posicao_a_nome_E = posicao_i_nome_E

posicao_i_cpf_S = [839, 336]
posicao_a_cpf_S = posicao_i_cpf_S

posicao_i_cpf_E = [590, 309]
posicao_a_cpf_E = posicao_i_cpf_E
posicao_i_lupa = [1690, 336]
posicao_a_lupa = posicao_i_lupa

for j in range(num):

    #copiar nome - SIGAMA
    pyautogui.moveTo(posicao_a_nome_S)
    pyautogui.tripleClick()
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.5)

    #localizar excel
    pyautogui.click(pyautogui.locateCenterOnScreen('excel_image.png', confidence = 0.8))
    time.sleep(0.5)

    #colar nome EXCEL
    pyautogui.moveTo(posicao_a_nome_E)
    pyautogui.click()
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'v')

    #localizar chrome
    pyautogui.click(pyautogui.locateCenterOnScreen('chrome_image.png', confidence = 0.8))

    #copiar cpf SIGAMA
    pyautogui.moveTo(posicao_a_cpf_S)
    pyautogui.doubleClick()
    pyautogui.hotkey('ctrl', 'c')

    #localizar excel
    pyautogui.click(pyautogui.locateCenterOnScreen('excel_image.png', confidence = 0.8))

    #colar cpf EXCEL
    pyautogui.moveTo(posicao_a_cpf_E)
    pyautogui.click()
    time.sleep(0.2)
    pyautogui.click()
    pyautogui.hotkey('_')
    time.sleep(0.2)
    pyautogui.hotkey('ctrl', 'v')

    #copiar nome e cpf EXCEL
    pyautogui.moveTo(posicao_a_nome_E)
    pyautogui.click()
    time.sleep(0.5)
    drag_x, drag_y = posicao_a_cpf_E
    pyautogui.dragTo(drag_x, drag_y, 0.2, button='left')
    pyautogui.hotkey('ctrl', 'c')
    texto = pyperclip.paste()

    texto_tratado = texto.replace("\t","").replace("\n", "").replace("\r", "")

    #abrir gerenciador de arquivos e criar nova pasta
    pyautogui.click(pyautogui.locateCenterOnScreen('arquivos_image.png', confidence = 0.8))
    time.sleep(0.5)
    pyautogui.click(pyautogui.locateCenterOnScreen('nova_pasta.png', confidence = 0.8))
    time.sleep(0.5)
    pyautogui.click(pyautogui.locateCenterOnScreen('pasta.png', confidence = 0.8))
    pyautogui.write(texto_tratado, interval=0.05)
    pyautogui.hotkey('Enter')

    #localizar chrome
    pyautogui.click(pyautogui.locateCenterOnScreen('chrome_image.png', confidence = 0.8))

    #localizar lupa
    pyautogui.moveTo(posicao_a_lupa)
    pyautogui.click()
    time.sleep(1)

    #localizar documentos
    pyautogui.moveTo(pyautogui.locateCenterOnScreen('anexo_image.png', confidence = 0.8))
    x1, y1 = pyautogui.locateCenterOnScreen('anexo_image.png', confidence = 0.8)

    #clicar no doc
    pyautogui.moveTo(x1, y1+27)

    for i in range(4):
        y1 = y1+23
        pyautogui.moveTo(x1, y1)
        time.sleep(0.5)
        pyautogui.click()
        time.sleep(0.2)

    pyautogui.click(pyautogui.locateCenterOnScreen('X_image.png', confidence = 0.8))

    #abrir gerenciador de arquivos e downloads
    pyautogui.click(pyautogui.locateCenterOnScreen('arquivos_image.png', confidence = 0.8))
    time.sleep(0.5)
    pyautogui.click(pyautogui.locateCenterOnScreen('download_image.png', confidence = 0.8))
    time.sleep(0.5)
    pyautogui.click(pyautogui.locateCenterOnScreen('hoje_image.png', confidence = 0.8))
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'c')
    pyautogui.click(pyautogui.locateCenterOnScreen('seta_voltar_image.png', confidence = 0.8))
    time.sleep(0.5)
    pyautogui.hotkey('Enter')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.click(pyautogui.locateCenterOnScreen('download_image.png', confidence = 0.8))
    time.sleep(0.5)
    pyautogui.click(pyautogui.locateCenterOnScreen('hoje_image.png', confidence = 0.8))
    time.sleep(1)
    pyautogui.hotkey('Delete')
    time.sleep(0.5)
    pyautogui.click(pyautogui.locateCenterOnScreen('seta_voltar_image.png', confidence = 0.8))
    time.sleep(0.5)
    pyautogui.click(pyautogui.locateCenterOnScreen('seta_voltar_image.png', confidence = 0.8))
    time.sleep(0.5)

    #localizar chrome
    pyautogui.click(pyautogui.locateCenterOnScreen('chrome_image.png', confidence = 0.8))

    posicao_a_nome_S[1] = posicao_a_nome_S[1] + 45

    posicao_a_nome_E[1] = posicao_a_nome_E[1] + 26

    posicao_a_cpf_S[1] = posicao_a_cpf_S[1] + 45

    posicao_a_cpf_E[1] = posicao_a_cpf_E[1] + 26

    posicao_a_lupa[1] = posicao_a_lupa[1] + 45
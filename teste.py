import pyautogui
import time

def copiar_colar():
    print("Você tem 5 segundos para posicionar o mouse na célula de origem...")
    time.sleep(5)
    origem = pyautogui.position()
    print(f"Posição da célula de origem capturada: {origem}")

    pyautogui.click(origem)
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(1)  # Tempo para garantir que copiou

    print("Agora posicione o mouse na célula de destino (5 segundos)...")
    time.sleep(5)
    destino = pyautogui.position()
    print(f"Posição da célula de destino capturada: {destino}")

    pyautogui.click(destino)
    pyautogui.hotkey('ctrl', 'v')
    print("Conteúdo colado com sucesso.")

copiar_colar()
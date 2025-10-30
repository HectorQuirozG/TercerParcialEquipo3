import subprocess
import pyautogui
import time
import logging
from datetime import datetime
from pathlib import Path

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3

def run_powershell(cmd):
    try:
        result = subprocess.run(["powershell", "-Command", cmd],
                                capture_output=True, text=True, timeout=10)
        return result.returncode, result.stdout.strip(), result.stderr.strip()
    except Exception as e:
        return 1, "", str(e)

def save_powershell(code, out, err):
    output = Path("out")
    output.mkdir(exist_ok=True)
    ts = datetime.now().strftime("%Y%m%dT%H%M%SZ")
    path = output / f"pws_{ts}.txt"
    with open(path, mode='w', encoding='utf-8') as f:
        f.write(out) if code == 0 else f.write(err)
    
def take_screenshot(name):
    out = Path("out")
    out.mkdir(exist_ok=True)
    ts = datetime.now().strftime("%Y%m%dT%H%M%SZ")
    path = out / f"{name}_{ts}.png"
    img = pyautogui.screenshot()
    img.save(path)
    return path

def fill_form(data, coords):
    take_screenshot("before")
    print("\nIniciando llenado de formulario en", end=' ')
    for i in range(3):
        print(f"{3 - i}...", end=' ')
        time.sleep(1)
    url = "https://shorturl.at/xAg0N"
    run_powershell(f"[System.Diagnostics.Process]::Start('{url}')")
    time.sleep(1)
    pyautogui.click(960, 540)
    for i in range(6):
        pyautogui.press("down")
        time.sleep(0.1)
    pyautogui.click(coords[0])
    pyautogui.click(coords[0])
    pyautogui.typewrite(data["fecha"])
    pyautogui.press("enter")
    time.sleep(1.5)
    pyautogui.click(coords[1])
    pyautogui.click(coords[1])
    for i in data["nombre"]:
        pyautogui.typewrite(i)
        pyautogui.press("enter")
    take_screenshot("during")
    time.sleep(0.5)
    pyautogui.click(coords[2])
    suma = sum(data["mats"])
    pyautogui.typewrite(str(suma))
    time.sleep(0.5)
    pyautogui.click(coords[3])
    pyautogui.click(960, 540)
    for i in range(6):
        pyautogui.press("down")
        time.sleep(0.1)
    pyautogui.click(coords[4])
    path = take_screenshot("after")
    logging.info(f"Capturas guardadas con éxito en {path}")

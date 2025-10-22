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

def fill_form(data, start_coords):
    take_screenshot("before")
    print("\nIniciando llenado de formulario en", end=' ')
    for i in range(3):
        print(f"{3 - i}...", end=' ')
        time.sleep(1)
    pyautogui.click(start_coords[0], start_coords[1])
    time.sleep(1)
    pyautogui.typewrite("notepad")
    time.sleep(1)
    pyautogui.press("enter")
    pyautogui.typewrite(data["nombre"])
    pyautogui.press("enter")
    pyautogui.typewrite(data["correo"])
    take_screenshot("during")
    pyautogui.press("enter")
    pyautogui.typewrite(data["equipo"])
    pyautogui.press("enter")
    time.sleep(1)
    path = take_screenshot("after")
    logging.info(f"Capturas guardadas con éxito en {path}")

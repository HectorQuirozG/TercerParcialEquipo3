from funcs import *
from datetime import datetime

def main():
    logging.basicConfig(filename="run.log", level=logging.INFO,
                        format="%(asctime)s %(levelname)s %(message)s", encoding="utf-8")
    logging.info("Inicio del examen")
    data = {
        "fecha": datetime.now().strftime('%d/%M/%Y'),
        "nombre": ["Hector Adrian Quiroz Gonzalez", "Valeria Rocha Solis",
                   "Hebert García Cantu", "Sofia del Carmen Chavez Reyna"],
        "mats": [2151398, 1746435, 2135105, 2143524]
    }
    coords = [(320, 212), (320, 365), (320, 600), (298, 732), (320,700)]
    code, out, err = run_powershell("(New-Object -ComObject Shell.Application).MinimizeAll()")
    save_powershell(code, out, err)
    print(f"PS output: {out}") if code == 0 else print(f"PS error: {err}")
    logging.info(f"PS code: {code}")
    logging.info(f"PS output: {out}") if code == 0 else logging.info(f"PS error: {err}")
    code, out, err = run_powershell("[System.Diagnostics.Process]::Start('https://www.fcfm.uanl.mx')")
    save_powershell(code, out, err)
    print(f"PS output: {out}") if code == 0 else print(f"PS error: {err}")
    logging.info(f"PS code: {code}")
    logging.info(f"PS output: {out}") if code == 0 else logging.info(f"PS error: {err}")
    fill_form(data, coords)
    logging.info("Fin del examen")
    
if __name__ == "__main__":
    main()

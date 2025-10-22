from funcs import *

def main():
    logging.basicConfig(filename="run.log", level=logging.INFO,
                        format="%(asctime)s %(levelname)s %(message)s", encoding="utf-8")
    logging.info("Inicio del examen")
    data = {
        "nombre": "Alumno Ejemplo",
        "correo": "ejemplo@correo.com",
        "equipo": "Equipo 3"
    }
    start_coords = (1, 767)
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
    fill_form(data, start_coords)
    logging.info("Fin del examen")
    
if __name__ == "__main__":
    main()

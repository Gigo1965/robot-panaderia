"""Ensayo del SaaS de demostracion EN LA PC, sobre una copia TEMPORAL.

Que hace:
  1. Copia datos_prueba/maestro_prueba.xlsx a una carpeta temporal de Windows.
  2. Arranca el SaaS apuntando a esa copia temporal (modo local: no usa Drive).
  3. Al cerrar (Ctrl + C) la copia temporal se borra sola.
Por eso NO hace falta "foto antes -> prueba -> restaurar": la copia local
datos_prueba/maestro_prueba.xlsx no se modifica, y Drive no se toca.

Uso (desde la carpeta ROBOT PANADERO):
    py -m saas_demo.ensayo
Luego abrir en el navegador:  http://127.0.0.1:5001
Los usuarios y claves del ensayo aparecen en la pantalla de ingreso (son de
prueba y solo existen en el ensayo; las claves reales del .env no se usan aqui).
"""
import atexit
import os
import shutil
import sys
import tempfile
from pathlib import Path

RAIZ_PROYECTO = Path(__file__).resolve().parent.parent
COPIA_LOCAL = RAIZ_PROYECTO / "datos_prueba" / "maestro_prueba.xlsx"
PUERTO = 5001


def preparar_copia_temporal():
    """Crea la copia temporal y deja el entorno en modo ENSAYO. Devuelve la ruta.

    Debe llamarse ANTES de importar saas_demo.app (asi nunca toca Drive).
    """
    if not COPIA_LOCAL.exists():
        sys.exit(f"No encontre la copia local: {COPIA_LOCAL}\n"
                 "El ensayo necesita datos_prueba\\maestro_prueba.xlsx (la misma de las otras pruebas).")
    carpeta = Path(tempfile.mkdtemp(prefix="saas_ensayo_"))
    destino = carpeta / "copia_ensayo.xlsx"
    shutil.copy2(COPIA_LOCAL, destino)
    atexit.register(shutil.rmtree, carpeta, ignore_errors=True)
    os.environ["SAAS_ENSAYO"] = "si"
    os.environ["ORIGEN_EXCEL"] = "local"
    os.environ["EXCEL_LOCAL_PATH"] = str(destino)
    return destino


def main():
    copia = preparar_copia_temporal()
    try:
        from .app import app
    except ModuleNotFoundError as error:
        if error.name == "flask":
            sys.exit("Falta la libreria Flask. En la Terminal (carpeta ROBOT PANADERO):\n"
                     "    py -m pip install -r whatsapp_bot/requirements.txt")
        raise
    print()
    print("ENSAYO del SaaS de demostracion (copia temporal; no toca Drive ni la copia local)")
    print(f"  Copia temporal: {copia}")
    print(f"  Abre en el navegador:  http://127.0.0.1:{PUERTO}")
    print("  Los usuarios de prueba estan en la pantalla de ingreso.")
    print("  Para terminar: Ctrl + C en esta Terminal.")
    print()
    app.run(host="127.0.0.1", port=PUERTO, debug=False)


if __name__ == "__main__":
    main()

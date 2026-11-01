"""Genera los Office Scripts autónomos (office-scripts/*.ts) a partir de office-scripts/src.

Office Scripts no admite import/export: cada script publicado contiene una copia
del bloque común (_comun.ts) a continuación de su propio código.
"""
import os

RAIZ = os.path.join(os.path.dirname(__file__), "..", "office-scripts")
SRC = os.path.join(RAIZ, "src")
NECESITA_COMUN = {"RegistrarCorreo", "EvaluarRelevancia", "ProcesarPendientes", "ActualizarDashboard",
                  "LimpiarDatosFicticios"}


def main():
    comun = open(os.path.join(SRC, "_comun.ts"), encoding="utf-8").read()
    for nombre in sorted(os.listdir(SRC)):
        if not nombre.endswith(".ts") or nombre.startswith("_"):
            continue
        base = nombre[:-3]
        cuerpo = open(os.path.join(SRC, nombre), encoding="utf-8").read()
        aviso = (f"// {base}.ts — ARCHIVO GENERADO desde office-scripts/src/{nombre}"
                 + (" + _comun.ts" if base in NECESITA_COMUN else "")
                 + ".\n// Copie TODO este archivo en Excel para la Web > Automatizar > Nuevo script.\n\n")
        salida = aviso + cuerpo + ("\n" + comun if base in NECESITA_COMUN else "")
        with open(os.path.join(RAIZ, nombre), "w", encoding="utf-8") as f:
            f.write(salida)
        print("generado", nombre)


if __name__ == "__main__":
    main()

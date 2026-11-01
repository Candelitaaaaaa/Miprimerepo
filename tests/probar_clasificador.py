"""Prueba el prompt y el esquema del clasificador contra la API de Claude con correos FICTICIOS.

No lee tu correo ni modifica el Excel. Úsalo para validar la clasificación antes
de conectar Power Automate.

    pip install anthropic
    export ANTHROPIC_API_KEY=...        # tu clave de https://console.anthropic.com
    python3 tests/probar_clasificador.py

Costo aproximado: 12 llamadas (menos de 1 USD con el modelo configurado).
"""
import json
import os

import anthropic

RAIZ = os.path.join(os.path.dirname(__file__), "..")
PA = os.path.join(RAIZ, "power-automate")


def main():
    plantilla = json.load(open(os.path.join(PA, "plantilla_cuerpo_claude.json"), encoding="utf-8"))
    prompt = open(os.path.join(PA, "prompt_sistema.txt"), encoding="utf-8").read()
    casos = json.load(open(os.path.join(os.path.dirname(__file__), "correos_prueba.json"), encoding="utf-8"))
    contexto = ("Empresas registradas: [DEMO] Minera Los Andes; [DEMO] Cobre Austral; [DEMO] Celulosa del Sur; "
                "[DEMO] Equipos Pacífico; [DEMO] Litio Atacama\n"
                "Postulaciones abiertas (ID | Empresa | Cargo | Estado):\n"
                "DEMO-0001 | [DEMO] Minera Los Andes | Memoria Planificación Mina 2027 | Entrevista con jefatura\n"
                "DEMO-0003 | [DEMO] Cobre Austral | Programa Memoristas 2027 | Evaluación online")
    client = anthropic.Anthropic()
    aciertos = 0
    for c in casos:
        mensaje = (f"Dirección: {c['direccion']}\nFecha: 2026-11-02T10:00\nRemitente: {c['de']}\n"
                   f"Para: {c.get('para', 'usuaria')}\nAsunto: {c['asunto']}\n\nTexto del correo:\n<<<\n{c['texto']}\n>>>\n\n"
                   f"{contexto}")
        r = client.messages.create(
            model=plantilla["model"],
            max_tokens=plantilla["max_tokens"],
            system=prompt,
            output_config=plantilla["output_config"],
            messages=[{"role": "user", "content": mensaje}],
        )
        if r.stop_reason != "end_turn":
            print(f"{c['id']}: stop_reason={r.stop_reason} (sin clasificación; el flujo usaría reglas)")
            continue
        texto = next(b.text for b in r.content if b.type == "text")
        d = json.loads(texto)
        ok = d["tipo_comunicacion"] == c["esperado"]
        aciertos += ok
        print(f"{'OK ' if ok else 'XX '} {c['id']}: {d['tipo_comunicacion']} (esperado {c['esperado']}) "
              f"conf={d['confianza']:.2f} empresa={d['empresa']} cargo={d['cargo']}")
    print(f"\nAciertos: {aciertos}/{len(casos)}")


if __name__ == "__main__":
    main()

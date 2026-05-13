"""
rfm_diario.py — Análisis RFM automático para Vexin Global Supply
Corre diario vía Programador de Tareas de Windows
"""

import os
import sys
import json
import subprocess
from datetime import datetime, timedelta
import pandas as pd
import anthropic

# ─── CONFIGURACIÓN ────────────────────────────────────────────────────────────
EXCEL_PATH = r"C:\Users\siste\Documents\20240419 Power BI Repot\14 lineas de producto en factura.xlsm"
REPO_DIR   = r"C:\Users\siste\rfm-vexin"
HTML_OUT   = os.path.join(REPO_DIR, "index.html")
LOG_FILE   = os.path.join(REPO_DIR, "rfm_log.txt")
# ──────────────────────────────────────────────────────────────────────────────

def log(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

def leer_excel():
    log("Leyendo Excel de Odoo...")
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=0, engine="openpyxl")
    except Exception:
        df = pd.read_excel(EXCEL_PATH, sheet_name=0)

    df.columns = df.columns.str.strip()
    log(f"Columnas encontradas: {list(df.columns)}")

    # Mapeo flexible de columnas
    mapeo = {
        "cliente":  ["cliente", "razón social", "partner", "nombre", "customer"],
        "fecha":    ["fecha de factura", "invoice date", "fecha", "date"],
        "total":    ["total", "amount", "monto", "importe", "amount_total"],
    }

    cols = {}
    for campo, opciones in mapeo.items():
        for col in df.columns:
            if col.lower().strip() in opciones:
                cols[campo] = col
                break
        if campo not in cols:
            # búsqueda parcial
            for col in df.columns:
                for op in opciones:
                    if op in col.lower():
                        cols[campo] = col
                        break
                if campo in cols:
                    break

    if len(cols) < 3:
        log(f"ERROR: No se encontraron todas las columnas. Encontradas: {cols}")
        sys.exit(1)

    log(f"Columnas mapeadas: {cols}")

    df = df[[cols["cliente"], cols["fecha"], cols["total"]]].copy()
    df.columns = ["cliente", "fecha", "total"]
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
    df["total"] = pd.to_numeric(df["total"], errors="coerce")
    df = df.dropna(subset=["cliente", "fecha", "total"])
    df = df[df["total"] > 0]

    # Últimos 12 meses
    corte = datetime.now() - timedelta(days=365)
    df = df[df["fecha"] >= corte]

    # Excluir montos de muestra
    df = df[df["total"] >= 5000]

    log(f"Registros válidos para análisis: {len(df)}")
    return df

def calcular_rfm(df):
    log("Calculando R, F, M...")
    hoy = datetime.now()

    rfm = df.groupby("cliente").agg(
        ultima_compra=("fecha", "max"),
        frecuencia=("fecha", "count"),
        monto=("total", "sum")
    ).reset_index()

    rfm["recencia_dias"] = (hoy - rfm["ultima_compra"]).dt.days

    n = len(rfm)
    if n >= 25:
        q = 5
    elif n >= 9:
        q = 3
    else:
        q = 2

    rfm["score_r"] = pd.qcut(rfm["recencia_dias"], q=q, labels=False, duplicates="drop")
    rfm["score_r"] = q - 1 - rfm["score_r"]  # invertir: menos días = score mayor
    rfm["score_f"] = pd.qcut(rfm["frecuencia"], q=q, labels=False, duplicates="drop")
    rfm["score_m"] = pd.qcut(rfm["monto"], q=q, labels=False, duplicates="drop")

    for col in ["score_r", "score_f", "score_m"]:
        rfm[col] = rfm[col] + 1  # escala 1-5

    rfm["score_total"] = rfm["score_r"] + rfm["score_f"] + rfm["score_m"]

    def segmento(row):
        r, f, m = row["score_r"], row["score_f"], row["score_m"]
        if r >= 4 and f >= 4 and m >= 4:
            return "Champion"
        elif f >= 4:
            return "Leal"
        elif m >= 4 and f <= 2:
            return "Alto valor potencial"
        elif r >= 4 and f == 1:
            return "Nuevo prometedor"
        elif r <= 2 and m >= 4:
            return "No puedo perderlos"
        elif r <= 2 and f >= 3:
            return "En riesgo"
        elif r <= 2 and f <= 2 and m <= 2:
            return "Hibernando"
        else:
            return "Necesita atención"

    rfm["segmento"] = rfm.apply(segmento, axis=1)
    rfm = rfm.sort_values("score_total", ascending=False)
    log(f"RFM calculado. Segmentos: {rfm['segmento'].value_counts().to_dict()}")
    return rfm

def generar_html_con_claude(rfm):
    log("Llamando a Claude para generar dashboard HTML...")

    resumen = rfm.groupby("segmento").agg(
        clientes=("cliente", "count"),
        facturacion=("monto", "sum"),
        ticket_prom=("monto", "mean")
    ).reset_index().to_dict(orient="records")

    top_clientes = rfm.head(50)[["cliente","segmento","recencia_dias","frecuencia","monto","score_r","score_f","score_m","score_total"]].to_dict(orient="records")

    fecha_hoy = datetime.now().strftime("%d/%m/%Y %H:%M")

    prompt = f"""Genera un dashboard HTML completo y profesional para vendedores de Vexin Global Supply con los resultados del análisis RFM del día de hoy ({fecha_hoy}).

DATOS DEL ANÁLISIS:

Resumen por segmento:
{json.dumps(resumen, ensure_ascii=False, indent=2, default=str)}

Top clientes (máximo 50):
{json.dumps(top_clientes, ensure_ascii=False, indent=2, default=str)}

REQUISITOS DEL HTML:
1. Archivo HTML único, autocontenido, sin dependencias externas
2. Diseño profesional con colores de Vexin: azul marino #1B2B5E y naranja #F47920
3. Encabezado con logo textual "VEXIN" y fecha de actualización
4. Tarjetas resumen por segmento con conteo, facturación y ticket promedio
5. Tabla de clientes con columnas: Cliente, Segmento, Días sin comprar, # Pedidos, Facturación, Score RFM
6. Código de colores por segmento:
   - Champion: verde #27ae60
   - En riesgo: rojo #e74c3c
   - No puedo perderlos: rojo oscuro #c0392b
   - Leal: azul #2980b9
   - Nuevo prometedor: verde claro #16a085
   - Hibernando: gris #7f8c8d
   - Alto valor potencial: naranja #f39c12
   - Necesita atención: gris azulado #95a5a6
7. Buscador de clientes por nombre
8. Filtro por segmento
9. Para cada segmento incluir una línea de acción recomendada al vendedor
10. Totalmente responsive para celular
11. Pie de página con "Actualizado automáticamente cada día · Vexin Global Supply"

Devuelve SOLO el código HTML completo, sin explicaciones, sin markdown, sin bloques de código."""

    client = anthropic.Anthropic()
    message = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=8000,
        messages=[{"role": "user", "content": prompt}]
    )

    html = message.content[0].text.strip()
    if html.startswith("```"):
        html = html.split("\n", 1)[1]
        html = html.rsplit("```", 1)[0]

    return html

def guardar_y_publicar(html):
    log("Guardando index.html...")
    with open(HTML_OUT, "w", encoding="utf-8") as f:
        f.write(html)

    log("Haciendo git push a GitHub Pages...")
    cmds = [
        ["git", "-C", REPO_DIR, "add", "index.html"],
        ["git", "-C", REPO_DIR, "add", "rfm_log.txt"],
        ["git", "-C", REPO_DIR, "commit", "-m", f"RFM actualizado {datetime.now().strftime('%Y-%m-%d %H:%M')}"],
        ["git", "-C", REPO_DIR, "push", "origin", "main"],
    ]
    for cmd in cmds:
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            log(f"Git warning: {result.stderr.strip()}")
        else:
            log(f"OK: {' '.join(cmd[2:])}")

def main():
    log("=" * 60)
    log("INICIANDO ANÁLISIS RFM DIARIO — VEXIN")
    log("=" * 60)

    try:
        df   = leer_excel()
        rfm  = calcular_rfm(df)
        html = generar_html_con_claude(rfm)
        guardar_y_publicar(html)
        log("✅ PROCESO COMPLETADO. Dashboard actualizado en GitHub Pages.")
    except Exception as e:
        log(f"❌ ERROR: {e}")
        import traceback
        log(traceback.format_exc())
        sys.exit(1)

if __name__ == "__main__":
    main()

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import tempfile
import os

# ── Constantes ─────────────────────────────────────────────────────
MESES = {1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
         7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"}

MESES_TEXTO = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
               "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
               "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}

PRECIO_BASE_USD = 37.8


def safe(val, default=0):
    if pd.notna(val) and isinstance(val, (int, float)):
        return val
    return default


# ── Lectura del .xlsb ─────────────────────────────────────────────
def extraer_mes_anno(df_proforma):
    texto = str(df_proforma.iloc[2, 2]).strip().lower()
    for nombre, num in MESES_TEXTO.items():
        if nombre in texto:
            anno = int(texto.split("de")[-1].strip())
            return anno, num
    raise ValueError(f"No se pudo extraer mes/año de: '{texto}'")


def extraer_servicio(row):
    ea = safe(row[36], 1)
    consumo = safe(row[32])
    cargo_ea = safe(row[38])
    pot = safe(row[37])
    cargo_pot = safe(row[39])
    return {
        "consumo_kwh": consumo,
        "energia_facturada": ea,
        "factor_ref": safe(row[18], 1),
        "dolar": safe(row[21]),
        "precio_base_usd": PRECIO_BASE_USD,
        "precio_energia_usd": safe(row[12]),
        "precio_energia_clp": cargo_ea / ea if ea else 0,
        "cargo_energia": cargo_ea,
        "pot_hp_kw": pot,
        "precio_potencia_kw": cargo_pot / pot if pot else 0,
        "cargo_potencia": cargo_pot,
        "tx_zonal_kwh": safe(row[26]),
        "tx_zonal_pesos": safe(row[26]) * consumo,
        "tx_nacional_kwh": safe(row[23]),
        "tx_nacional_pesos": safe(row[23]) * consumo,
        "exenciones_kwh": safe(row[25]),
        "exenciones_pesos": safe(row[25]) * consumo,
        "sscc_kwh": safe(row[24]),
        "sscc_pesos": safe(row[24]) * consumo,
        "csp_kwh": safe(row[31]),
        "csp_pesos": safe(row[43]),
        "sobrecostos_mt": safe(row[46]),
        "sobrecosto_pd": safe(row[48]),
        "compensacion_pe": safe(row[50]),
        "costo_oportunidad_rh": 0,
        "sobrecosto_rh": 0,
        "serv_complementarios": safe(row[47]),
        "reliquidacion": 0,
        "liquidacion_peaje_cen": safe(row[49]),
        "pago_ernc": 0,
    }


def extraer_peajes_dx(df_peajes):
    for r in range(len(df_peajes) - 1, 1, -1):
        cobrado = df_peajes.iloc[r, 37]
        if pd.notna(cobrado) and str(cobrado).strip().upper() == "NO":
            return parsear_peaje(df_peajes.iloc[r])
    return parsear_peaje(df_peajes.iloc[len(df_peajes) - 1])


def parsear_peaje(row):
    energia = safe(row[31])
    hp = safe(row[34])
    dda = safe(row[32])
    pot_fc = safe(row[33])
    estab = safe(row[18])
    dda_p = safe(row[19])
    comp_p = safe(row[20])
    hp_p = safe(row[21])
    return {
        "energia": energia, "hp": hp, "dda_max": dda, "pot_fact_comp": pot_fc,
        "cargo_fijo": safe(row[17]),
        "cargo_energia_kwh": 0, "cargo_energia_pesos": 0,
        "cargo_dda_max_kw": dda_p / dda if dda else 0, "cargo_dda_max_pesos": dda_p,
        "cargo_compra_pot_kw": comp_p / pot_fc if pot_fc else 0, "cargo_compra_pot_pesos": comp_p,
        "cargo_hp_kw": hp_p / hp if hp else 0, "cargo_hp_pesos": hp_p,
        "estab_kwh": estab / energia if energia else 0, "estab_pesos": estab,
    }


def leer_facturacion(ruta):
    xls = pd.ExcelFile(ruta, engine="pyxlsb")
    bd = pd.read_excel(xls, sheet_name="BD Clientes", header=None)
    pq = pd.read_excel(xls, sheet_name="PROFORMA-Q", header=None)
    pl = pd.read_excel(xls, sheet_name="PROFORMA-L", header=None)
    pdx = pd.read_excel(xls, sheet_name="PeajesDx_Limache", header=None)

    anno, mes_num = extraer_mes_anno(pq)
    quilpue = extraer_servicio(bd.iloc[2])
    limache = extraer_servicio(bd.iloc[3])

    quilpue["reliquidacion"] = safe(pq.iloc[40, 5]) if len(pq) > 40 else 0
    limache["reliquidacion"] = safe(pl.iloc[40, 5]) if len(pl) > 40 else 0

    return {
        "anno": anno, "mes_num": mes_num, "mes_str": MESES[mes_num],
        "quilpue": quilpue, "limache": limache,
        "peajes_dx": extraer_peajes_dx(pdx),
    }


# ── Escritura en el .xlsx ─────────────────────────────────────────
def encontrar_fila(ws, anno, mes):
    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=1).value == anno and ws.cell(row=row, column=2).value == mes:
            return row
    for row in range(2, ws.max_row + 2):
        if ws.cell(row=row, column=1).value is None:
            return row
    return ws.max_row + 1


def escribir_seat(ws, r, d):
    vals = [
        d["anno"], d["mes"], d["consumo_kwh"], d["energia_facturada"],
        d["factor_ref"], d["dolar"], d["precio_base_usd"], d["precio_energia_usd"],
        d["precio_energia_clp"], d["cargo_energia"], d["pot_hp_kw"],
        d["precio_potencia_kw"], d["cargo_potencia"],
        d["tx_zonal_kwh"], d["tx_zonal_pesos"],
        d["tx_nacional_kwh"], d["tx_nacional_pesos"],
        d["exenciones_kwh"], d["exenciones_pesos"],
        d["sscc_kwh"], d["sscc_pesos"],
        d["csp_kwh"], d["csp_pesos"],
        d["sobrecostos_mt"], d["sobrecosto_pd"], d["compensacion_pe"],
        d["costo_oportunidad_rh"], d["sobrecosto_rh"],
        d["serv_complementarios"], d["reliquidacion"],
        d["liquidacion_peaje_cen"], d["pago_ernc"],
    ]
    for col, val in enumerate(vals, start=1):
        ws.cell(row=r, column=col, value=val)
    ws.cell(row=r, column=33, value=f"=J{r}+M{r}+O{r}+Q{r}+S{r}+U{r}+W{r}+X{r}+Y{r}+Z{r}+AA{r}+AB{r}+AC{r}+AD{r}+AE{r}+AF{r}")
    ws.cell(row=r, column=34, value=f"=(AG{r}-W{r})*0.19")
    ws.cell(row=r, column=35, value=f"=AG{r}+AH{r}")


def escribir_limache(ws, r, d, p):
    vals = [
        d["anno"], d["mes"], d["consumo_kwh"], d["energia_facturada"],
        d["factor_ref"], d["dolar"], d["precio_base_usd"], d["precio_energia_usd"],
        d["precio_energia_clp"], d["cargo_energia"], d["pot_hp_kw"],
        d["precio_potencia_kw"], d["cargo_potencia"],
        d["tx_zonal_kwh"], d["tx_zonal_pesos"],
        d["tx_nacional_kwh"], d["tx_nacional_pesos"],
        d["exenciones_kwh"], d["exenciones_pesos"],
        d["sscc_kwh"], d["sscc_pesos"],
        d["csp_kwh"], d["csp_pesos"],
        d["sobrecostos_mt"], d["sobrecosto_pd"], d["compensacion_pe"],
        d["costo_oportunidad_rh"], d["sobrecosto_rh"],
        d["serv_complementarios"], d["reliquidacion"],
        d["liquidacion_peaje_cen"], d["pago_ernc"],
        p["energia"], p["hp"], p["dda_max"], p["pot_fact_comp"],
        p["cargo_fijo"], p["cargo_energia_kwh"], p["cargo_energia_pesos"],
        p["cargo_dda_max_kw"], p["cargo_dda_max_pesos"],
        p["cargo_compra_pot_kw"], p["cargo_compra_pot_pesos"],
        p["cargo_hp_kw"], p["cargo_hp_pesos"],
        p["estab_kwh"], p["estab_pesos"],
    ]
    for col, val in enumerate(vals, start=1):
        ws.cell(row=r, column=col, value=val)
    ws.cell(row=r, column=48, value=f"=J{r}+M{r}+O{r}+Q{r}+S{r}+U{r}+W{r}+X{r}+Y{r}+Z{r}+AA{r}+AB{r}+AC{r}+AD{r}+AE{r}+AF{r}+AK{r}+AM{r}+AO{r}+AQ{r}+AS{r}+AU{r}")
    ws.cell(row=r, column=49, value=f"=AV{r}*0.19")
    ws.cell(row=r, column=50, value=f"=AV{r}+AW{r}")


def procesar(fact_bytes, proy_bytes):
    """Procesa ambos archivos y retorna el xlsx resultante como bytes."""
    # Guardar facturación en archivo temporal (pyxlsb necesita ruta)
    with tempfile.NamedTemporaryFile(suffix=".xlsb", delete=False) as tmp:
        tmp.write(fact_bytes)
        tmp_path = tmp.name

    try:
        datos = leer_facturacion(tmp_path)
    finally:
        os.unlink(tmp_path)

    anno, mes = datos["anno"], datos["mes_str"]

    wb = load_workbook(BytesIO(proy_bytes))

    ws_q = wb["Fac. Ea. SEAT"]
    fila_q = encontrar_fila(ws_q, anno, mes)
    escribir_seat(ws_q, fila_q, {**datos["quilpue"], "anno": anno, "mes": mes})

    ws_l = wb["Fac. Ea. Limache"]
    fila_l = encontrar_fila(ws_l, anno, mes)
    escribir_limache(ws_l, fila_l, {**datos["limache"], "anno": anno, "mes": mes}, datos["peajes_dx"])

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output, datos, fila_q, fila_l


# ── Interfaz Streamlit ─────────────────────────────────────────────
st.set_page_config(page_title="Facturación → Proyección", page_icon="⚡")
st.title("⚡ Facturación → Proyección Energía")
st.write("Sube los dos archivos y descarga la proyección actualizada.")

col1, col2 = st.columns(2)

with col1:
    fact_file = st.file_uploader(
        "📄 Archivo de Facturación (.xlsb)",
        type=["xlsb"],
        help="Ej: 2026_01_Facturacion_Contrato_ENELEFE_Valpo_15m.xlsb"
    )

with col2:
    proy_file = st.file_uploader(
        "📊 Archivo de Proyección (.xlsx)",
        type=["xlsx"],
        help="Ej: Proyección_Energía.xlsx"
    )

if fact_file and proy_file:
    if st.button("🔄 Procesar", type="primary", use_container_width=True):
        with st.spinner("Procesando..."):
            try:
                resultado, datos, fila_q, fila_l = procesar(
                    fact_file.read(), proy_file.read()
                )

                st.success(f"✅ Período: **{datos['mes_str']} {datos['anno']}**")

                c1, c2 = st.columns(2)
                with c1:
                    st.metric("QUILPUE", f"{datos['quilpue']['energia_facturada']:,.0f} kWh")
                    st.caption(f"→ Fac. Ea. SEAT fila {fila_q}")
                with c2:
                    st.metric("LIMACHE", f"{datos['limache']['energia_facturada']:,.0f} kWh")
                    st.caption(f"→ Fac. Ea. Limache fila {fila_l}")

                nombre_salida = proy_file.name.replace(".xlsx", "_actualizado.xlsx")
                st.download_button(
                    label="⬇️ Descargar Proyección Actualizada",
                    data=resultado,
                    file_name=nombre_salida,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )

            except Exception as e:
                st.error(f"Error: {e}")
else:
    st.info("👆 Sube ambos archivos para comenzar.")

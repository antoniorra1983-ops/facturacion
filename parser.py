"""Lee y extrae datos del archivo de Facturación Contrato ENELEFE (.xlsb)."""
import pandas as pd

from .constants import (
    BD, BD_ROWS, PDX,
    PROFORMA_RELIQ_ROW, PROFORMA_RELIQ_COL,
    MESES, MESES_TEXTO, PRECIO_BASE_USD_MWH,
)


def _safe(val, default=0):
    """Retorna el valor si es numérico, sino el default."""
    if pd.notna(val) and isinstance(val, (int, float)):
        return val
    return default


def _extraer_mes_anno(df_proforma):
    """Extrae año y mes desde la celda 'enero de 2026' en la PROFORMA."""
    texto = str(df_proforma.iloc[2, 2]).strip().lower()
    for nombre, num in MESES_TEXTO.items():
        if nombre in texto:
            anno = int(texto.split("de")[-1].strip())
            return anno, num
    raise ValueError(f"No se pudo extraer mes/año de: '{texto}'")


def _extraer_servicio(row, consumo_col=BD.ENE_LEIDO):
    """Extrae un diccionario con los datos de un punto de suministro desde BD Clientes."""
    ea_facturada = _safe(row[BD.ENERGIA_FACTURADA], 1)
    consumo = _safe(row[consumo_col])
    cargo_ea = _safe(row[BD.CARGO_ENERGIA])
    pot_hp = _safe(row[BD.DHP_FACT_KW])
    cargo_pot = _safe(row[BD.CARGO_POTENCIA])

    return {
        "consumo_kwh": consumo,
        "energia_facturada": ea_facturada,
        "factor_ref": _safe(row[BD.FACTOR_REF], 1),
        "dolar": _safe(row[BD.TIPO_CAMBIO]),
        "precio_base_usd": PRECIO_BASE_USD_MWH,
        "precio_energia_usd": _safe(row[BD.PEBASE_INDEX]),
        "precio_energia_clp": cargo_ea / ea_facturada if ea_facturada else 0,
        "cargo_energia": cargo_ea,
        "pot_hp_kw": pot_hp,
        "precio_potencia_kw": cargo_pot / pot_hp if pot_hp else 0,
        "cargo_potencia": cargo_pot,
        "tx_zonal_kwh": _safe(row[BD.TX_ZONAL_KWH]),
        "tx_zonal_pesos": _safe(row[BD.TX_ZONAL_KWH]) * consumo,
        "tx_nacional_kwh": _safe(row[BD.TX_NACIONAL_KWH]),
        "tx_nacional_pesos": _safe(row[BD.TX_NACIONAL_KWH]) * consumo,
        "exenciones_kwh": _safe(row[BD.EXENCIONES_KWH]),
        "exenciones_pesos": _safe(row[BD.EXENCIONES_KWH]) * consumo,
        "sscc_kwh": _safe(row[BD.TX_SSCC_KWH]),
        "sscc_pesos": _safe(row[BD.TX_SSCC_KWH]) * consumo,
        "csp_kwh": _safe(row[BD.CSP_TOTAL_KWH]),
        "csp_pesos": _safe(row[BD.CSP_PESOS]),
        "sobrecostos_mt": _safe(row[BD.SOBRECOSTOS_MT]),
        "sobrecosto_pd": _safe(row[BD.SOBRECOSTO_PD]),
        "compensacion_pe": _safe(row[BD.PRECIO_ESTAB]),
        "costo_oportunidad_rh": 0,
        "sobrecosto_rh": 0,
        "serv_complementarios": _safe(row[BD.SERV_COMPLEMENTARIOS]),
        "reliquidacion": 0,  # se completa después desde PROFORMA
        "liquidacion_peaje_cen": _safe(row[BD.PEAJE_TX_NAC]),
        "pago_ernc": 0,
    }


def _extraer_reliquidacion(df_proforma):
    """Extrae el monto de reliquidación desde una hoja PROFORMA."""
    if len(df_proforma) <= PROFORMA_RELIQ_ROW:
        return 0
    val = df_proforma.iloc[PROFORMA_RELIQ_ROW, PROFORMA_RELIQ_COL]
    return _safe(val)


def _extraer_peajes_dx(df_peajes):
    """Extrae los cargos de peaje distribución de Limache (última fila no cobrada)."""
    for r in range(len(df_peajes) - 1, 1, -1):
        cobrado = df_peajes.iloc[r, PDX.COBRADO]
        if pd.notna(cobrado) and str(cobrado).strip().upper() == "NO":
            return _parsear_fila_peaje(df_peajes.iloc[r])
    # Si todas están cobradas, usar la última fila de datos
    return _parsear_fila_peaje(df_peajes.iloc[len(df_peajes) - 1])


def _parsear_fila_peaje(row):
    """Convierte una fila de PeajesDx_Limache en diccionario."""
    energia = _safe(row[PDX.ENERGIA])
    hp = _safe(row[PDX.HP])
    dda_max = _safe(row[PDX.DDA_MAX])
    pot_fc = _safe(row[PDX.POT_FACT_COMP])
    cargo_fijo = _safe(row[PDX.CARGO_FIJO])
    estab_pesos = _safe(row[PDX.PEAJE_ENERGIA])
    dda_pesos = _safe(row[PDX.PEAJE_DDA_MAX])
    compra_pesos = _safe(row[PDX.PEAJE_COMPRA_POT])
    hp_pesos = _safe(row[PDX.PEAJE_HP])

    return {
        "energia": energia,
        "hp": hp,
        "dda_max": dda_max,
        "pot_fact_comp": pot_fc,
        "cargo_fijo": cargo_fijo,
        "cargo_energia_kwh": 0,
        "cargo_energia_pesos": 0,
        "cargo_dda_max_kw": dda_pesos / dda_max if dda_max else 0,
        "cargo_dda_max_pesos": dda_pesos,
        "cargo_compra_pot_kw": compra_pesos / pot_fc if pot_fc else 0,
        "cargo_compra_pot_pesos": compra_pesos,
        "cargo_hp_kw": hp_pesos / hp if hp else 0,
        "cargo_hp_pesos": hp_pesos,
        "estab_kwh": estab_pesos / energia if energia else 0,
        "estab_pesos": estab_pesos,
    }


def leer_facturacion(ruta_xlsb: str) -> dict:
    """
    Lee el archivo de facturación y retorna un dict con todos los datos
    necesarios para escribir en la Proyección Energía.

    Returns:
        dict con claves: anno, mes_num, mes_str, quilpue, limache, peajes_dx
    """
    xls = pd.ExcelFile(ruta_xlsb, engine="pyxlsb")

    bd = pd.read_excel(xls, sheet_name="BD Clientes", header=None)
    proforma_q = pd.read_excel(xls, sheet_name="PROFORMA-Q", header=None)
    proforma_l = pd.read_excel(xls, sheet_name="PROFORMA-L", header=None)
    peajes_dx = pd.read_excel(xls, sheet_name="PeajesDx_Limache", header=None)

    anno, mes_num = _extraer_mes_anno(proforma_q)

    quilpue = _extraer_servicio(bd.iloc[BD_ROWS.QUILPUE_ACTUAL])
    quilpue["reliquidacion"] = _extraer_reliquidacion(proforma_q)

    limache = _extraer_servicio(bd.iloc[BD_ROWS.LIMACHE_ACTUAL])
    limache["reliquidacion"] = _extraer_reliquidacion(proforma_l)

    peaje = _extraer_peajes_dx(peajes_dx)

    return {
        "anno": anno,
        "mes_num": mes_num,
        "mes_str": MESES[mes_num],
        "quilpue": quilpue,
        "limache": limache,
        "peajes_dx": peaje,
    }

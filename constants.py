"""Constantes y mapeos de columnas."""

MESES = {
    1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic",
}

MESES_TEXTO = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
}

PRECIO_BASE_USD_MWH = 37.8  # Precio base contrato

# ── BD Clientes: índices de columna (0-based) ──────────────────────
class BD:
    SIGLA = 7
    PEBASE_INDEX = 12          # Precio energía actual USD/MWh
    FACTOR_REF = 18            # Factor Referenciación de Precios
    TIPO_CAMBIO = 21           # Dólar observado
    PE_CLP_KWH = 22           # Precio energía $/kWh
    TX_NACIONAL_KWH = 23      # Transmisión Nacional $/kWh
    TX_SSCC_KWH = 24          # SSCC $/kWh
    EXENCIONES_KWH = 25       # Exenciones $/kWh
    TX_ZONAL_KWH = 26         # Transmisión Zonal $/kWh
    CSP_TOTAL_KWH = 31        # CSP Total $/kWh
    ENE_LEIDO = 32            # Consumo kWh (medido)
    ENERGIA_FACTURADA = 36    # Energía facturada kWh
    DHP_FACT_KW = 37          # Demanda HP facturada kW
    CARGO_ENERGIA = 38        # Cargo por Energía $
    CARGO_POTENCIA = 39       # Cargo por demanda HP $
    CSP_PESOS = 43            # Cargo por Servicio Público $
    SOBRECOSTOS_MT = 46       # Min. Técnicos (Sobrecostos MT)
    SERV_COMPLEMENTARIOS = 47 # Servicios Complementarios
    SOBRECOSTO_PD = 48        # Sobrecostos Part. y Detenciones
    PEAJE_TX_NAC = 49         # Peaje Tx Nacional (Cliente Indiv.)
    PRECIO_ESTAB = 50         # Precio Estabilizado (PMG)
    RESERVA_HIDRICA = 51      # Reserva Hídrica
    RELIQUIDACION = 52        # Reliquidación
    PEAJES_DX = 53            # Peajes Distribución
    TOTAL_SIN_IVA = 55
    TOTAL_AFECTO = 56
    TOTAL_EXENTO = 57
    IVA = 58
    TOTAL_PAGAR = 59

# ── BD Clientes: filas (0-based) ───────────────────────────────────
class BD_ROWS:
    QUILPUE_ACTUAL = 2
    LIMACHE_ACTUAL = 3
    QUILPUE_ANTERIOR = 7
    LIMACHE_ANTERIOR = 8
    QUILPUE_RELIQ = 9
    LIMACHE_RELIQ = 10

# ── PROFORMA: fila de reliquidación (0-based) ─────────────────────
PROFORMA_RELIQ_ROW = 40
PROFORMA_RELIQ_COL = 5

# ── PeajesDx_Limache: columnas (0-based) ──────────────────────────
class PDX:
    CARGO_FIJO = 17
    PEAJE_ENERGIA = 18       # → Estabilización Tarifas $
    PEAJE_DDA_MAX = 19       # → Cargo dda máxima potencia $
    PEAJE_COMPRA_POT = 20    # → Cargo compra potencia $
    PEAJE_HP = 21            # → Cargo dda máx HP $
    ENERGIA = 31             # Energía kWh (medición Dx)
    DDA_MAX = 32             # Demanda máxima kW
    POT_FACT_COMP = 33       # Potencia facturada compensada kW
    HP = 34                  # Horas punta kW
    COBRADO = 37             # "SI" / "NO"

# ── Fac. Ea. SEAT: columnas destino (1-based, openpyxl) ───────────
class SEAT:
    AÑO = 1
    MES = 2
    CONSUMO = 3
    EA_FACTURADA = 4
    FACTOR_REF = 5
    DOLAR = 6
    PRECIO_BASE = 7
    PRECIO_EA_USD = 8
    PRECIO_EA_CLP = 9
    CARGO_EA = 10
    POT_HP_KW = 11
    PRECIO_POT_KW = 12
    CARGO_POT = 13
    TX_ZONAL_KWH = 14
    TX_ZONAL = 15
    TX_NAC_KWH = 16
    TX_NAC = 17
    EXENC_KWH = 18
    EXENC = 19
    SSCC_KWH = 20
    SSCC = 21
    CSP_KWH = 22
    CSP = 23
    SOBRE_MT = 24
    SOBRE_PD = 25
    COMP_PE = 26
    COSTO_RH = 27
    SOBRE_RH = 28
    SERV_COMP = 29
    RELIQ = 30
    PEAJE_CEN = 31
    PAGO_ERNC = 32
    TOTAL = 33
    IVA = 34
    TOTAL_IVA = 35

# ── Fac. Ea. Limache: columnas destino (1-based, openpyxl) ────────
class LIM:
    # Columnas 1-32: igual que SEAT
    ENERGIA_DX = 33
    HP_DX = 34
    DDA_MAX_DX = 35
    POT_FACT_COMP = 36
    CARGO_FIJO = 37
    CARGO_EA_KWH = 38
    CARGO_EA = 39
    CARGO_DDA_KW = 40
    CARGO_DDA = 41
    CARGO_COMPRA_KW = 42
    CARGO_COMPRA = 43
    CARGO_HP_KW = 44
    CARGO_HP = 45
    ESTAB_KWH = 46
    ESTAB = 47
    TOTAL = 48
    IVA = 49
    TOTAL_IVA = 50

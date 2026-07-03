#!/usr/bin/env python3
"""
Transfiere datos del archivo de Facturación Contrato ENELEFE (.xlsb)
al formato del archivo Proyección Energía (.xlsx).

Uso:
    python -m src.main <facturacion.xlsb> <proyeccion.xlsx> [-o salida.xlsx]
"""
import argparse
import sys
from pathlib import Path

from .parser import leer_facturacion
from .writer import escribir_proyeccion


def main():
    parser = argparse.ArgumentParser(
        description="Transfiere datos de Facturación ENELEFE al formato Proyección Energía."
    )
    parser.add_argument("facturacion", help="Archivo de facturación (.xlsb)")
    parser.add_argument("proyeccion", help="Archivo de proyección energía (.xlsx)")
    parser.add_argument(
        "-o", "--output",
        help="Archivo de salida (default: <proyeccion>_actualizado.xlsx)",
        default=None,
    )
    args = parser.parse_args()

    ruta_fact = Path(args.facturacion)
    ruta_proy = Path(args.proyeccion)

    if not ruta_fact.exists():
        print(f"ERROR: No se encontró {ruta_fact}", file=sys.stderr)
        sys.exit(1)
    if not ruta_proy.exists():
        print(f"ERROR: No se encontró {ruta_proy}", file=sys.stderr)
        sys.exit(1)

    salida = args.output or str(ruta_proy).replace(".xlsx", "_actualizado.xlsx")

    print(f"Leyendo facturación: {ruta_fact}")
    datos = leer_facturacion(str(ruta_fact))
    print(f"  Período: {datos['mes_str']} {datos['anno']}")
    print(f"  QUILPUE  - Energía: {datos['quilpue']['energia_facturada']:,.0f} kWh")
    print(f"  LIMACHE  - Energía: {datos['limache']['energia_facturada']:,.0f} kWh")

    print(f"\nEscribiendo en: {salida}")
    resultado = escribir_proyeccion(str(ruta_proy), datos, salida)

    for msg in resultado["mensajes"]:
        print(f"  {msg}")

    print(f"\n✓ Archivo guardado: {salida}")


if __name__ == "__main__":
    main()

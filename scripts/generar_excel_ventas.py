"""Genera Excel_de_Ventas_CEA.xlsx: la plantilla operativa del Tablero Comercial CEA.

Reproducible: correr `python scripts/generar_excel_ventas.py` desde la raíz del
repo regenera el archivo desde cero. Los nombres de hoja y columna DEBEN
coincidir con lo que leen src/sheets.py y src/etl.py — si cambias algo aquí,
cambia también esos dos archivos.
"""
from __future__ import annotations

import datetime as dt
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from src.constants import (  # noqa: E402
    APROBADORES,
    CANALES,
    ESTRATEGIAS,
    ETAPAS,
    KPIS_MANUALES,
    KR_IDS,
    MOTIVOS_PERDIDA,
    PLAZAS,
    PROGRAMAS,
    RESPONSABLES,
    TIPOS_DESCUENTO,
    VIGENCIA_DESCUENTO,
)

AUTOR_COMENTARIOS = "Obsidian"
SALIDA = REPO_ROOT / "Excel_de_Ventas_CEA.xlsx"

GRIS_CALCULADO = PatternFill("solid", fgColor="EAEAEA")
GRIS_EJEMPLO = PatternFill("solid", fgColor="D9D9D9")
VERDE_ALUMNO = PatternFill("solid", fgColor="C6E0B4")
GRIS_PERDIDO = PatternFill("solid", fgColor="D9D9D9")
ROJO_TOPE = PatternFill("solid", fgColor="F8CBAD")

FUENTE_ENCABEZADO = Font(bold=True)
FMT_FECHA = "DD/MM/AAAA"
FMT_MXN = '#,##0 "MXN"'


def nueva_hoja(wb: Workbook, nombre: str, *, con_grid: bool = True):
    ws = wb.create_sheet(nombre)
    if not con_grid:
        ws.sheet_view.showGridLines = False
    return ws


def encabezados(ws, columnas: list[str]) -> None:
    for i, nombre in enumerate(columnas, start=1):
        celda = ws.cell(row=1, column=i, value=nombre)
        celda.font = FUENTE_ENCABEZADO
    ws.freeze_panes = "A2"


def comentario(ws, celda: str, texto: str) -> None:
    c = Comment(texto, AUTOR_COMENTARIOS)
    c.width = 320
    c.height = 140
    ws[celda].comment = c


def agregar_validacion_lista(wb: Workbook, ws, columna_letra: str, nombre_rango: str, fila_inicio: int, fila_fin: int) -> None:
    dv = DataValidation(type="list", formula1=f"={nombre_rango}", allow_blank=True, showDropDown=False)
    ws.add_data_validation(dv)
    dv.add(f"{columna_letra}{fila_inicio}:{columna_letra}{fila_fin}")


def agregar_validacion_fecha(ws, columna_letra: str, fila_inicio: int, fila_fin: int) -> None:
    dv = DataValidation(type="date", operator="greaterThan", formula1="DATE(2020,1,1)", allow_blank=True)
    dv.error = "Escribe una fecha válida (DD/MM/AAAA)."
    ws.add_data_validation(dv)
    dv.add(f"{columna_letra}{fila_inicio}:{columna_letra}{fila_fin}")
    for fila in range(fila_inicio, fila_fin + 1):
        ws[f"{columna_letra}{fila}"].number_format = FMT_FECHA


def agregar_validacion_decimal(ws, columna_letra: str, fila_inicio: int, fila_fin: int, minimo: float, maximo: float) -> None:
    dv = DataValidation(type="decimal", operator="between", formula1=str(minimo), formula2=str(maximo), allow_blank=True)
    dv.error = f"Debe estar entre {minimo} y {maximo}."
    ws.add_data_validation(dv)
    dv.add(f"{columna_letra}{fila_inicio}:{columna_letra}{fila_fin}")


def pintar_filas_ejemplo(ws, fila_inicio: int, fila_fin: int, num_columnas: int) -> None:
    for fila in range(fila_inicio, fila_fin + 1):
        for col in range(1, num_columnas + 1):
            ws.cell(row=fila, column=col).fill = GRIS_EJEMPLO


def congelar_y_filtrar(ws, num_columnas: int) -> None:
    ultima_col = get_column_letter(num_columnas)
    ws.auto_filter.ref = f"A1:{ultima_col}1"


def ajustar_anchos(ws, anchos: dict[str, int]) -> None:
    for columna, ancho in anchos.items():
        ws.column_dimensions[columna].width = ancho


def construir_config(wb: Workbook) -> dict[str, tuple[str, int]]:
    """Crea la hoja Config con una columna por lista, y un rango con nombre
    por cada una (usado por las validaciones de lista de las demás hojas).
    Regresa {nombre_lista: (columna_letra, cantidad_de_valores)}."""
    ws = nueva_hoja(wb, "Config")
    comentario(ws, "A1", "Solo Obsidian edita esta hoja. Agregar aquí = aparece en las listas de todas las hojas.")

    listas = {
        "etapas": ETAPAS,
        "canales": CANALES,
        "plazas": PLAZAS,
        "programas": PROGRAMAS,
        "responsables": RESPONSABLES,
        "aprobadores": APROBADORES,
        "tipos_descuento": TIPOS_DESCUENTO,
        "motivos_perdida": MOTIVOS_PERDIDA,
        "vigencia_descuento": VIGENCIA_DESCUENTO,
        "estrategias": ESTRATEGIAS + ["Total"],
        "kr_ids": KR_IDS,
        "kpis_manuales": KPIS_MANUALES,
    }

    rangos: dict[str, tuple[str, int]] = {}
    for i, (nombre_lista, valores) in enumerate(listas.items(), start=1):
        col_letra = get_column_letter(i)
        ws.cell(row=1, column=i, value=nombre_lista).font = FUENTE_ENCABEZADO
        for j, valor in enumerate(valores, start=2):
            ws.cell(row=j, column=i, value=valor)
        rango_nombre = f"lista_{nombre_lista}"
        wb.defined_names[rango_nombre] = DefinedName(
            rango_nombre, attr_text=f"'Config'!${col_letra}$2:${col_letra}${len(valores) + 1}"
        )
        rangos[nombre_lista] = (col_letra, len(valores))

    # Tabla programa -> mensualidad plena, para el formato condicional de Descuentos.
    with open(REPO_ROOT / "data" / "precios_cea.json", encoding="utf-8") as f:
        precios = json.load(f)["programas"]
    mensualidad_por_programa = {
        "Lic. Administración": next(p["mensualidad_plena"] for p in precios if p["programa"].startswith("Licenciatura")),
        "Lic. Contaduría": next(p["mensualidad_plena"] for p in precios if p["programa"].startswith("Licenciatura")),
        "Lic. Derecho": next(p["mensualidad_plena"] for p in precios if p["programa"].startswith("Licenciatura")),
        "Bachillerato": next(p["mensualidad_plena"] for p in precios if p["programa"] == "Bachillerato Sabatino" and p["plaza"] == "Puebla"),
        "Maestría Derecho Penal": next(p["mensualidad_plena"] for p in precios if p["programa"].startswith("Maestría")),
    }
    col_prog = get_column_letter(len(listas) + 2)
    col_mens = get_column_letter(len(listas) + 3)
    ws.cell(row=1, column=len(listas) + 2, value="programa_precio").font = FUENTE_ENCABEZADO
    ws.cell(row=1, column=len(listas) + 3, value="mensualidad_plena").font = FUENTE_ENCABEZADO
    for j, (programa, monto) in enumerate(mensualidad_por_programa.items(), start=2):
        ws.cell(row=j, column=len(listas) + 2, value=programa)
        ws.cell(row=j, column=len(listas) + 3, value=monto)
    wb.defined_names["tabla_mensualidad_plena"] = DefinedName(
        "tabla_mensualidad_plena",
        attr_text=f"'Config'!${col_prog}$2:${col_mens}${len(mensualidad_por_programa) + 1}",
    )

    ajustar_anchos(ws, {get_column_letter(i): 22 for i in range(1, len(listas) + 4)})
    return rangos


def construir_instrucciones(wb: Workbook) -> None:
    ws = nueva_hoja(wb, "INSTRUCCIONES", con_grid=False)
    ws.column_dimensions["A"].width = 100

    lineas = [
        ("titulo", "Excel de Ventas CEA — cómo se llena"),
        ("texto", ""),
        ("sub", "¿Qué es esto?"),
        ("texto", "Este archivo alimenta el Tablero Comercial CEA que se revisa cada lunes. "
                   "Cada hoja corresponde a una parte del ritual semanal: prospección, avance de "
                   "metas, alumnos activos, descuentos y los KPIs que no salen de ningún otro lado."),
        ("texto", ""),
        ("sub", "Regla de oro"),
        ("texto", "Si no está aquí, no existió. El tablero solo cree lo que esté capturado antes "
                   "del lunes 8:00 am."),
        ("texto", ""),
        ("sub", "Quién llena qué"),
        ("texto", "• Asesoras → hoja Prospectos."),
        ("texto", "• Control Escolar → hoja Alumnos_Activos y parte de KPI_Manual (NPS, bajas)."),
        ("texto", "• Paola → hoja Alumnos_Activos (ingreso/facturable/cobrado), Descuentos_Otorgados y cobranza."),
        ("texto", "• C. Gutiérrez → KPI_Manual (reseñas, prospectos por canal de pauta)."),
        ("texto", "• Obsidian → hojas Metas y OKR_Avance, en vivo durante el comité de los lunes."),
        ("texto", ""),
        ("sub", "Los 3 errores que rompen el tablero"),
        ("texto", "1. Escribir un valor a mano en vez de elegir de la lista desplegable."),
        ("texto", "2. Borrar o reciclar el id de un prospecto."),
        ("texto", "3. Capturar el teléfono en la columna de notas."),
        ("texto", ""),
        ("sub", "Nunca tocar la hoja Eventos"),
        ("texto", "Se llena sola con el Apps Script instalado (ver apps_script/registro_eventos.gs). "
                   "Escribir ahí a mano rompe las métricas de velocidad del embudo."),
        ("texto", ""),
        ("sub", "Aviso de privacidad"),
        ("texto", "Los datos de contacto (teléfono, notas) solo se usan para dar seguimiento a la "
                   "inscripción, conforme a la LFPDPPP. El tablero público nunca muestra nombres ni "
                   "teléfonos — esas columnas se descartan antes de llegar a cualquier gráfica."),
        ("texto", ""),
        ("sub", "Antes de arrancar"),
        ("texto", "Borra las 3 filas que empiezan con \"EJEMPLO —\" en cada hoja de captura. Son "
                   "solo para mostrar cómo se llena, no son datos reales."),
    ]

    fila = 1
    for tipo, texto in lineas:
        celda = ws.cell(row=fila, column=1, value=texto)
        if tipo == "titulo":
            celda.font = Font(bold=True, size=16)
        elif tipo == "sub":
            celda.font = Font(bold=True, size=12)
        else:
            celda.alignment = Alignment(wrap_text=True, vertical="top")
        fila += 1


def main() -> None:
    wb = Workbook()
    wb.remove(wb.active)
    construir_instrucciones(wb)
    construir_config(wb)
    wb.save(SALIDA)
    print(f"Escrito {SALIDA} con hojas: {wb.sheetnames}")


if __name__ == "__main__":
    main()

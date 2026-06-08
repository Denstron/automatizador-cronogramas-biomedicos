"""
AUTOMATIZADOR DE CRONOGRAMAS BIOMÉDICOS V1
==========================================
ESE Hospital San Rafael Yolombó
Autor: Daniel España
Fecha: Junio 2026

Qué hace este script:
1. Lee el cronograma Excel
2. Detecta equipos vencidos (mes de mantenimiento ya pasó)
3. Detecta equipos próximos (mantenimiento en los próximos 30 días)
4. Genera reporte Excel con colores
"""

import pandas as pd
from datetime import datetime
import os

# ============================================================
# CONFIGURACIÓN
# ============================================================

ARCHIVO_ENTRADA = "Cronograma_de_mantenimiento_ESE_Hospital_Yolombo.xlsx"
ARCHIVO_SALIDA  = "Reporte_Mantenimiento_V1.xlsx"

# Mes actual (Junio = 6)
HOY = datetime.today()
MES_ACTUAL = HOY.month   # 6
AÑO_ACTUAL = HOY.year    # 2026

# Mapeo de nombres de meses a números
MESES = {
    "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4,
    "MAYO": 5, "JUNIO": 6, "JULIO": 7, "AGOSTO": 8,
    "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12
}

# ============================================================
# PASO 1: LEER EL ARCHIVO
# ============================================================

print("=" * 60)
print("AUTOMATIZADOR DE CRONOGRAMAS BIOMÉDICOS V1")
print("=" * 60)
print(f"\nFecha actual: {HOY.strftime('%d/%m/%Y')}")
print(f"Leyendo archivo: {ARCHIVO_ENTRADA}\n")

# Leer Excel - fila 0 son los encabezados
df = pd.read_excel(ARCHIVO_ENTRADA, header=0)

# Renombrar columnas para facilitar el trabajo
df.columns = [
    "ITEM", "DESCRIPCION", "UBICACION", "N_INVENTARIO",
    "FABRICANTE", "MARCA", "MODELO", "SERIE",
    "CLASIFICACION_RIESGO", "REGISTRO_INVIMA",
    "FECHA_COMPRA", "PERIODICIDAD", "FECHA_MANTENIMIENTO"
]

# Eliminar filas completamente vacías
df = df.dropna(subset=["DESCRIPCION"])

# Reemplazar NR por vacío
df.replace("NR", "", inplace=True)

print(f"Total equipos encontrados: {len(df)}")

# ============================================================
# PASO 2: FUNCIÓN PARA EXTRAER MESES DE LA COLUMNA FECHA
# ============================================================

def extraer_meses(texto_fecha):
    """
    Convierte texto como "ABRIL / OCTUBRE" en lista [4, 10]

    Parámetros:
    - texto_fecha: string con los meses separados por "/"

    Retorna:
    - Lista de números de mes [4, 10]
    - Lista vacía si no hay datos
    """
    if pd.isna(texto_fecha) or str(texto_fecha).strip() == "":
        return []

    # Separar por "/" y limpiar espacios
    partes = str(texto_fecha).upper().split("/")

    meses_encontrados = []
    for parte in partes:
        parte = parte.strip()
        if parte in MESES:
            meses_encontrados.append(MESES[parte])

    return meses_encontrados

# ============================================================
# PASO 3: FUNCIÓN PARA CALCULAR ESTADO DEL EQUIPO
# ============================================================

def calcular_estado(meses_mantenimiento):
    """
    Dado los meses de mantenimiento, calcula el estado del equipo.

    Estados posibles:
    - VENCIDO: El mes de mantenimiento ya pasó este año
    - PRÓXIMO: El mantenimiento es este mes o el próximo
    - OK: El mantenimiento está a más de 1 mes
    - SIN FECHA: No hay datos de fecha

    Parámetros:
    - meses_mantenimiento: lista de números [4, 10]

    Retorna:
    - Estado (string)
    - Próximo mes de mantenimiento (número)
    """
    if not meses_mantenimiento:
        return "SIN FECHA", None

    # Ordenar meses
    meses = sorted(meses_mantenimiento)

    # Buscar próximo mes de mantenimiento
    proximo_mes = None
    for mes in meses:
        if mes >= MES_ACTUAL:
            proximo_mes = mes
            break

    # Si no hay próximo mes este año, el último mes ya pasó
    if proximo_mes is None:
        ultimo_mes = meses[-1]
        meses_vencido = MES_ACTUAL - ultimo_mes
        return f"VENCIDO ({meses_vencido} mes(es))", ultimo_mes

    # Calcular diferencia con mes actual
    diferencia = proximo_mes - MES_ACTUAL

    if diferencia == 0:
        return "🔴 ESTE MES", proximo_mes
    elif diferencia == 1:
        return "🟡 PRÓXIMO MES", proximo_mes
    elif diferencia <= 2:
        return "🟠 EN 2 MESES", proximo_mes
    else:
        return "🟢 OK", proximo_mes

# ============================================================
# PASO 4: APLICAR A TODOS LOS EQUIPOS
# ============================================================

print("\nAnalizando equipos...\n")

# Extraer meses de cada equipo
df["MESES_PARSED"] = df["FECHA_MANTENIMIENTO"].apply(extraer_meses)

# Calcular estado y próximo mes
resultados = df["MESES_PARSED"].apply(calcular_estado)
df["ESTADO"]      = resultados.apply(lambda x: x[0])
df["PROXIMO_MES"] = resultados.apply(lambda x: x[1])

# Convertir número de mes a nombre
MESES_INVERSO = {v: k for k, v in MESES.items()}
df["PROXIMO_MES_NOMBRE"] = df["PROXIMO_MES"].apply(
    lambda x: MESES_INVERSO.get(x, "") if x else ""
)

# ============================================================
# PASO 5: MOSTRAR RESUMEN EN CONSOLA
# ============================================================

print("=" * 60)
print("RESUMEN DE ESTADO")
print("=" * 60)

conteo = df["ESTADO"].value_counts()
for estado, cantidad in conteo.items():
    print(f"  {estado}: {cantidad} equipos")

print(f"\nTotal: {len(df)} equipos")

# Mostrar equipos críticos
print("\n" + "=" * 60)
print("⚠️  EQUIPOS VENCIDOS Y PRÓXIMOS:")
print("=" * 60)

criticos = df[df["ESTADO"].str.contains("VENCIDO|ESTE MES|PRÓXIMO", na=False)]

if len(criticos) == 0:
    print("✅ No hay equipos vencidos o próximos.")
else:
    for _, row in criticos.iterrows():
        print(f"\n  📋 {row['DESCRIPCION']}")
        print(f"     Ubicación  : {row['UBICACION']}")
        print(f"     Inventario : {row['N_INVENTARIO']}")
        print(f"     Estado     : {row['ESTADO']}")
        print(f"     Fecha prog : {row['FECHA_MANTENIMIENTO']}")

# ============================================================
# PASO 6: GENERAR REPORTE EXCEL
# ============================================================

print(f"\nGenerando reporte: {ARCHIVO_SALIDA}...")

# Columnas para el reporte final
columnas_reporte = [
    "ITEM", "DESCRIPCION", "UBICACION", "N_INVENTARIO",
    "PERIODICIDAD", "FECHA_MANTENIMIENTO",
    "PROXIMO_MES_NOMBRE", "ESTADO"
]

df_reporte = df[columnas_reporte].copy()

# Guardar como Excel
with pd.ExcelWriter(ARCHIVO_SALIDA, engine="openpyxl") as writer:
    df_reporte.to_excel(writer, sheet_name="CRONOGRAMA", index=False)

print(f"✅ Reporte guardado: {ARCHIVO_SALIDA}")

print("\n" + "=" * 60)
print("✅ AUTOMATIZADOR V1 COMPLETADO")
print("=" * 60)

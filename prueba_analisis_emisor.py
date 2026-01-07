# -*- coding: utf-8 -*-
"""
Created on Tue Jan  6 14:46:54 2026

@author:David Riveros Morales
"""

# Prueba Técnica – Analista de Riesgo de Crédito
# Autor: David Riveros Morales


import pandas as pd
import os
import calendar
from datetime import date


# Ruta archivos Excel

ruta_data = r"C:\Users\user\Desktop\David\Skandia\prueba_analista_riesgo\data"


# Función para construir la fecha

def fecha_desde_archivo(nombre_archivo):
    nombre = nombre_archivo.replace(".xlsx", "")
    mes_txt, anio = nombre.split("_")

    meses = {
        "ene": 1, "feb": 2, "mar": 3, "abr": 4,
        "may": 5, "jun": 6, "jul": 7, "ago": 8,
        "sep": 9, "oct": 10, "nov": 11, "dic": 12
    }

    mes = meses[mes_txt.lower()]
    anio = int(anio)

    ultimo_dia = calendar.monthrange(anio, mes)[1]

    #return datetime(anio, mes, ultimo_dia)
    return date(anio, mes, ultimo_dia)


# Función principal:
# Lee todos los archivos Excel de la ruta, normaliza la
# estructura contable y consolida la información

def cargar_y_consolidar_archivos(ruta_data):

    archivos = [
        f for f in os.listdir(ruta_data)
        if f.endswith(".xlsx")
    ]

    lista_df = []

    for archivo in archivos:
        ruta_archivo = os.path.join(ruta_data, archivo)
        fecha = fecha_desde_archivo(archivo)

        # Lectura del archivo base
        df = pd.read_excel(
            ruta_archivo,
            header=None,
            skiprows=10,
            usecols="A:BB"
        )

        # Renombrar columnas relevantes
        df = df.rename(columns={
            0: "Rubro",
            1: "Subrubro",
            2: "Detalle",
            3: "Total_Entidades",
            53: "LULO_BANK_SA"
        })

        df = df[
            ["Rubro", "Subrubro", "Detalle",
             "Total_Entidades", "LULO_BANK_SA"]
        ]

        # Eliminar filas completamente vacías
        df = df.dropna(how="all")

        # Normalización jerárquica

        # Rubro siempre se propaga
        df["Rubro"] = df["Rubro"].ffill()

        # Subrubro:
        # - Se propaga solo cuando pertenece al mismo rubro
        # - Se mantiene vacío cuando representa totales
        df["Subrubro"] = df.groupby("Rubro")["Subrubro"].ffill()

        # Detalle:
        # - Nunca se propaga
        # - Si está vacío, representa total de rubro o subrubro

        # Limpieza semántica
        # Eliminar filas sin ningún concepto (seguridad)
        df = df[
            df[["Rubro", "Subrubro", "Detalle"]].notna().any(axis=1)
        ]

        # Eliminar textos informativos no contables
        df = df[
            ~df["Detalle"].astype(str).str.contains(
                "Nombre de la entidad|Cifras con corte",
                case=False,
                na=False
            )
        ]

        # Conversión a valores numéricos
        df["Total_Entidades"] = pd.to_numeric(
            df["Total_Entidades"], errors="coerce"
        )
        df["LULO_BANK_SA"] = pd.to_numeric(
            df["LULO_BANK_SA"], errors="coerce"
        )

        # Asignar fecha
        df["Fecha"] = fecha

        lista_df.append(df)

    base_consolidada = pd.concat(lista_df, ignore_index=True)

    return base_consolidada


# Ejecución
base_consolidada = cargar_y_consolidar_archivos(ruta_data)

# Exportar resultado
base_consolidada.to_excel(
    r"C:\Users\user\Desktop\David\Skandia\prueba_analista_riesgo\resultado\base_consolidada.xlsx",
    index=False
)

#### Graficos
# Activos Totales
import matplotlib.pyplot as plt

df_activos_totales = base_consolidada[
    (base_consolidada["Rubro"].str.upper() == "ACTIVOS") &
    (base_consolidada["Subrubro"].isna()) &
    (base_consolidada["Detalle"].isna())
]

activos_totales = (
    df_activos_totales
    .groupby("Fecha")[["Total_Entidades", "LULO_BANK_SA"]]
    .sum()
)

plt.figure()
plt.plot(
    activos_totales.index,
    activos_totales["Total_Entidades"],
    label="Sistema"
)
plt.plot(
    activos_totales.index,
    activos_totales["LULO_BANK_SA"],
    label="LULO BANK"
)

plt.title("Activos Totales (Rubro ACTIVOS) – LULO BANK vs Sistema")
plt.xlabel("Fecha")
plt.ylabel("Valor")
plt.legend()
plt.tight_layout()
plt.show()

# Endeudamiento y apalancamiento

# Filtrar solo cuentas generales por rubro
df_totales_rubro = base_consolidada[
    (base_consolidada["Subrubro"].isna()) &
    (base_consolidada["Detalle"].isna()) &
    (base_consolidada["Rubro"].isin(["ACTIVOS", "PASIVOS"]))
]

# Agrupar por Fecha y Rubro (sistema y LULO)
df_rubro_anual = (
    df_totales_rubro
    .groupby(["Fecha", "Rubro"])[["Total_Entidades", "LULO_BANK_SA"]]
    .sum()
    .reset_index()
)

# Separar activos
df_activos = df_rubro_anual[df_rubro_anual["Rubro"] == "ACTIVOS"] \
    .rename(columns={
        "Total_Entidades": "Activos_Sistema",
        "LULO_BANK_SA": "Activos_Lulo"
    }) \
    .drop(columns="Rubro")

# Separar pasivos
df_pasivos = df_rubro_anual[df_rubro_anual["Rubro"] == "PASIVOS"] \
    .rename(columns={
        "Total_Entidades": "Pasivos_Sistema",
        "LULO_BANK_SA": "Pasivos_Lulo"
    }) \
    .drop(columns="Rubro")

# Unir activos y pasivos por Fecha
df_endeudamiento = pd.merge(
    df_activos,
    df_pasivos,
    on="Fecha",
    how="inner"
)

# Calcular ratios de endeudamiento
df_endeudamiento["Endeudamiento_Sistema"] = (
    df_endeudamiento["Pasivos_Sistema"] /
    df_endeudamiento["Activos_Sistema"]
)

df_endeudamiento["Endeudamiento_Lulo"] = (
    df_endeudamiento["Pasivos_Lulo"] /
    df_endeudamiento["Activos_Lulo"]
)

df_endeudamiento

# Grafica endeudamiento
plt.figure()

plt.plot(
    df_endeudamiento["Fecha"],
    df_endeudamiento["Endeudamiento_Sistema"],
    label="Sistema"
)

plt.plot(
    df_endeudamiento["Fecha"],
    df_endeudamiento["Endeudamiento_Lulo"],
    label="LULO BANK"
)

plt.title("Endeudamiento (Pasivo / Activo)\nSistema vs LULO BANK")
plt.xlabel("Fecha")
plt.ylabel("Pasivo / Activo")
plt.legend()
plt.tight_layout()
plt.show()

# ROE


# Filtrar solo cuentas generales para ROE
df_roe_base = base_consolidada[
    (base_consolidada["Subrubro"].isna()) &
    (base_consolidada["Detalle"].isna()) &
    (base_consolidada["Rubro"].isin([
        "GANANCIAS (EXCEDENTES) Y PÉRDIDAS",
        "PATRIMONIO"
    ]))
]

# Agrupar por Fecha y Rubro
df_roe_anual = (
    df_roe_base
    .groupby(["Fecha", "Rubro"])[["Total_Entidades", "LULO_BANK_SA"]]
    .sum()
    .reset_index()
)

# Separar ganancias / pérdidas
df_resultado = df_roe_anual[
    df_roe_anual["Rubro"] == "GANANCIAS (EXCEDENTES) Y PÉRDIDAS"
].rename(columns={
    "Total_Entidades": "Resultado_Sistema",
    "LULO_BANK_SA": "Resultado_Lulo"
}).drop(columns="Rubro")

# Separar patrimonio
df_patrimonio = df_roe_anual[
    df_roe_anual["Rubro"] == "PATRIMONIO"
].rename(columns={
    "Total_Entidades": "Patrimonio_Sistema",
    "LULO_BANK_SA": "Patrimonio_Lulo"
}).drop(columns="Rubro")

# Unir resultados y patrimonio
df_roe = pd.merge(
    df_resultado,
    df_patrimonio,
    on="Fecha",
    how="inner"
)

# Calcular ROE
df_roe["ROE_Sistema"] = (
    df_roe["Resultado_Sistema"] /
    df_roe["Patrimonio_Sistema"]
)

df_roe["ROE_Lulo"] = (
    df_roe["Resultado_Lulo"] /
    df_roe["Patrimonio_Lulo"]
)

df_roe

## Gráfico ROE


plt.figure()

plt.plot(
    df_roe["Fecha"],
    df_roe["ROE_Sistema"],
    label="Sistema"
)

plt.plot(
    df_roe["Fecha"],
    df_roe["ROE_Lulo"],
    label="LULO BANK"
)

plt.title("ROE – Sistema vs LULO BANK")
plt.xlabel("Fecha")
plt.ylabel("ROE")
plt.legend()
plt.tight_layout()
plt.show()



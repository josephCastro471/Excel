"""
Módulo para procesar archivos Excel de ventas.
Autor: Joseph Castro
Fecha: 2026-07-05

Este módulo automatiza el proceso de:
1. Leer un Excel con ventas (Producto, Cantidad, Precio Unitario)
2. Calcular el total por producto
3. Generar un resumen con TOTAL GENERAL
"""

import pandas as pd
from pathlib import Path
import logging

# ============================================================
# CONFIGURACIÓN DE LOGS
# ============================================================
# Los logs ayudan a saber qué está pasando mientras el programa ejecuta
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# ============================================================
# FUNCIÓN 1: LEER ARCHIVO EXCEL
# ============================================================
def leer_excel(ruta_archivo):
    """
    Lee un archivo Excel y devuelve un DataFrame de pandas.
    
    Args:
        ruta_archivo (str): Ruta del archivo Excel.
    
    Returns:
        pd.DataFrame: Datos del archivo.
    
    Raises:
        FileNotFoundError: Si el archivo no existe.
        ValueError: Si las columnas no son las esperadas.
    
    Ejemplo:
        >>> df = leer_excel("data/ventas.xlsx")
        >>> print(len(df))  # Muestra cuántas filas tiene
    """
    # Verificar que el archivo existe
    if not Path(ruta_archivo).exists():
        raise FileNotFoundError(f"❌ El archivo {ruta_archivo} no existe.")
    
    # Leer el archivo Excel
    df = pd.read_excel(ruta_archivo)
    
    # Verificar que tiene las columnas esperadas
    columnas_esperadas = ['Producto', 'Cantidad', 'Precio Unitario']
    columnas_faltantes = [col for col in columnas_esperadas if col not in df.columns]
    
    if columnas_faltantes:
        raise ValueError(f"❌ Faltan columnas: {columnas_faltantes}")
    
    logging.info(f"✅ Archivo leído: {ruta_archivo} - {len(df)} filas encontradas")
    return df

# ============================================================
# FUNCIÓN 2: CALCULAR TOTALES POR PRODUCTO
# ============================================================
def calcular_totales_por_producto(df):
    """
    Calcula el total de ventas por producto.
    
    Args:
        df (pd.DataFrame): DataFrame con columnas Producto, Cantidad, Precio Unitario.
    
    Returns:
        pd.DataFrame: DataFrame con Producto y Total Ventas.
    
    Ejemplo:
        >>> df = pd.DataFrame({'Producto': ['A', 'B'], 'Cantidad': [2, 3], 'Precio Unitario': [10, 20]})
        >>> resultado = calcular_totales_por_producto(df)
        >>> print(resultado)
          Producto  Total Ventas
        0        A            20
        1        B            60
    """
    # Crear columna "Total" multiplicando Cantidad por Precio Unitario
    df['Total'] = df['Cantidad'] * df['Precio Unitario']
    
    # Agrupar por producto y sumar los totales
    resumen = df.groupby('Producto', as_index=False)['Total'].sum()
    
    # Renombrar la columna para que sea más clara
    resumen = resumen.rename(columns={'Total': 'Total Ventas'})
    
    logging.info(f"✅ Totales calculados para {len(resumen)} productos")
    return resumen

# ============================================================
# FUNCIÓN 3: AGREGAR FILA DE TOTAL GENERAL
# ============================================================
def agregar_total_general(df_resumen):
    """
    Agrega una fila con el total general al final del resumen.
    
    Args:
        df_resumen (pd.DataFrame): DataFrame con Producto y Total Ventas.
    
    Returns:
        pd.DataFrame: DataFrame con fila de total general al final.
    
    Ejemplo:
        >>> df = pd.DataFrame({'Producto': ['A', 'B'], 'Total Ventas': [100, 200]})
        >>> resultado = agregar_total_general(df)
        >>> print(resultado)
          Producto  Total Ventas
        0        A           100
        1        B           200
        2  TOTAL GENERAL       300
    """
    # Calcular la suma de todos los totales
    total_general = df_resumen['Total Ventas'].sum()
    
    # Crear una fila con el total general
    fila_total = pd.DataFrame({
        'Producto': ['TOTAL GENERAL'],
        'Total Ventas': [total_general]
    })
    
    # Concatenar el resumen con la fila de total
    df_con_total = pd.concat([df_resumen, fila_total], ignore_index=True)
    
    logging.info(f"✅ Total general agregado: {total_general}")
    return df_con_total

# ============================================================
# FUNCIÓN 4: GUARDAR ARCHIVO EXCEL
# ============================================================
def guardar_excel(df, ruta_salida):
    """
    Guarda un DataFrame en un archivo Excel.
    
    Args:
        df (pd.DataFrame): DataFrame a guardar.
        ruta_salida (str): Ruta donde guardar el archivo.
    
    Returns:
        bool: True si se guardó correctamente.
    
    Ejemplo:
        >>> df = pd.DataFrame({'A': [1, 2], 'B': [3, 4]})
        >>> guardar_excel(df, "data/salida.xlsx")
        True
    """
    # Crear la carpeta si no existe
    Path(ruta_salida).parent.mkdir(parents=True, exist_ok=True)
    
    # Guardar el DataFrame en Excel
    df.to_excel(ruta_salida, index=False)
    
    logging.info(f"✅ Archivo guardado: {ruta_salida}")
    return True

# ============================================================
# FUNCIÓN PRINCIPAL: PROCESAR EXCEL COMPLETO
# ============================================================
def procesar_excel_ventas(ruta_entrada, ruta_salida):
    """
    Procesa el archivo Excel de ventas y genera el resumen.
    
    Esta función orquesta todo el proceso:
    1. Leer el archivo de entrada
    2. Calcular totales por producto
    3. Agregar total general
    4. Guardar el resultado
    
    Args:
        ruta_entrada (str): Ruta del archivo Excel de entrada.
        ruta_salida (str): Ruta donde guardar el resumen.
    
    Returns:
        dict: Diccionario con el resultado del procesamiento.
    
    Ejemplo:
        >>> resultado = procesar_excel_ventas("data/ventas.xlsx", "data/resumen_ventas.xlsx")
        >>> print(resultado['total_general'])
        2775
    """
    logging.info(f"🚀 Iniciando procesamiento de {ruta_entrada}")
    
    try:
        # PASO 1: Leer archivo
        df = leer_excel(ruta_entrada)
        
        # PASO 2: Calcular totales por producto
        df_resumen = calcular_totales_por_producto(df)
        
        # PASO 3: Agregar total general
        df_con_total = agregar_total_general(df_resumen)
        
        # PASO 4: Guardar resultado
        guardar_excel(df_con_total, ruta_salida)
        
        # Devolver información del proceso
        return {
            'exito': True,
            'mensaje': f"✅ Procesamiento exitoso. Archivo guardado en {ruta_salida}",
            'total_general': df_con_total['Total Ventas'].iloc[-1],
            'productos_procesados': len(df_resumen),
            'filas_procesadas': len(df)
        }
        
    except Exception as e:
        logging.error(f"❌ Error en el procesamiento: {e}")
        raise

# ============================================================
# EJECUCIÓN DIRECTA (cuando se corre el archivo)
# ============================================================
if __name__ == "__main__":
    """
    Esta parte se ejecuta solo cuando corremos este archivo directamente.
    Si el archivo es importado (import procesar_excel), esta parte NO se ejecuta.
    """
    
    # Definir las rutas de los archivos
    ARCHIVO_ENTRADA = "data/ventas.xlsx"
    ARCHIVO_SALIDA = "data/resumen_ventas.xlsx"
    
    print("=" * 60)
    print("📊 PROCESADOR DE VENTAS - EXCEL")
    print("=" * 60)
    
    try:
        # Ejecutar el procesamiento
        resultado = procesar_excel_ventas(ARCHIVO_ENTRADA, ARCHIVO_SALIDA)
        
        # Mostrar resultados
        print("\n✅ ¡PROCESAMIENTO EXITOSO!")
        print(f"📊 Total general de ventas: ${resultado['total_general']:,.2f}")
        print(f"📝 Productos procesados: {resultado['productos_procesados']}")
        print(f"📄 Filas leídas: {resultado['filas_procesadas']}")
        print(f"💾 Archivo guardado en: {ARCHIVO_SALIDA}")
        
    except FileNotFoundError as e:
        print(f"\n❌ {e}")
        print("💡 Sugerencia: Asegúrate de que el archivo 'ventas.xlsx' existe en la carpeta 'data/'")
        
    except ValueError as e:
        print(f"\n❌ {e}")
        print("💡 Sugerencia: Verifica que el archivo tenga las columnas: 'Producto', 'Cantidad', 'Precio Unitario'")
        
    except Exception as e:
        print(f"\n❌ Error inesperado: {e}")
    
    print("\n" + "=" * 60)
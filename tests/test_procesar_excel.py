"""
Pruebas unitarias para el módulo procesar_excel.py
Autor: Joseph Castro
Fecha: 2026-07-05

Estas pruebas verifican que todas las funciones funcionen correctamente.
"""

import pytest
import pandas as pd
import os
import tempfile
from pathlib import Path
from src.procesar_excel import (
    leer_excel,
    calcular_totales_por_producto,
    agregar_total_general,
    guardar_excel,
    procesar_excel_ventas
)

# ============================================================
# DATOS DE PRUEBA (se usan en varias pruebas)
# ============================================================
DATOS_PRUEBA = pd.DataFrame({
    'Producto': ['Laptop', 'Mouse', 'Teclado', 'Laptop', 'Mouse'],
    'Cantidad': [2, 10, 3, 1, 5],
    'Precio Unitario': [800, 15, 45, 800, 15]
})

# Resultados esperados:
# Laptop: 2*800 + 1*800 = 2400
# Mouse: 10*15 + 5*15 = 225
# Teclado: 3*45 = 135
# Total general: 2400 + 225 + 135 = 2760

# ============================================================
# PRUEBAS PARA leer_excel()
# ============================================================
def test_leer_excel():
    """Prueba que la función leer_excel lea correctamente un archivo Excel."""
    # Crear un archivo Excel temporal
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp_path = tmp.name
        DATOS_PRUEBA.to_excel(tmp_path, index=False)
    
    try:
        # Leer el archivo
        df = leer_excel(tmp_path)
        
        # Verificar que se leyó correctamente
        assert len(df) == 5, "Debería leer 5 filas"
        assert list(df.columns) == ['Producto', 'Cantidad', 'Precio Unitario'], "Columnas incorrectas"
        assert df.iloc[0]['Producto'] == 'Laptop', "Primer producto debería ser Laptop"
        
    finally:
        # Eliminar el archivo temporal
        if Path(tmp_path).exists():
            os.unlink(tmp_path)

def test_leer_excel_archivo_no_existente():
    """Prueba que leer_excel lance error cuando el archivo no existe."""
    with pytest.raises(FileNotFoundError) as excinfo:
        leer_excel("archivo_que_no_existe.xlsx")
    
    assert "no existe" in str(excinfo.value), "El mensaje de error debería indicar que no existe"

def test_leer_excel_columnas_incorrectas():
    """Prueba que leer_excel lance error cuando faltan columnas."""
    # Crear Excel con columnas incorrectas
    df_incorrecto = pd.DataFrame({
        'Producto': ['A', 'B'],
        'Cantidad': [1, 2],
        'Precio': [10, 20]  # Debería ser 'Precio Unitario'
    })
    
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp_path = tmp.name
        df_incorrecto.to_excel(tmp_path, index=False)
    
    try:
        with pytest.raises(ValueError) as excinfo:
            leer_excel(tmp_path)
        
        assert "Faltan columnas" in str(excinfo.value), "Debería indicar que faltan columnas"
        
    finally:
        if Path(tmp_path).exists():
            os.unlink(tmp_path)

# ============================================================
# PRUEBAS PARA calcular_totales_por_producto()
# ============================================================
def test_calcular_totales_por_producto():
    """Prueba el cálculo de totales por producto."""
    # Crear copia de los datos de prueba
    df = DATOS_PRUEBA.copy()
    
    # Calcular totales
    resultado = calcular_totales_por_producto(df)
    
    # Verificar que el resultado tiene las columnas correctas
    assert list(resultado.columns) == ['Producto', 'Total Ventas'], "Columnas incorrectas"
    
    # Verificar los totales
    # Laptop: 2*800 + 1*800 = 2400
    total_laptop = resultado[resultado['Producto'] == 'Laptop']['Total Ventas'].iloc[0]
    assert total_laptop == 9999, f"Laptop debería ser 2400, pero es {total_laptop}"
    
    # Mouse: 10*15 + 5*15 = 225
    total_mouse = resultado[resultado['Producto'] == 'Mouse']['Total Ventas'].iloc[0]
    assert total_mouse == 225, f"Mouse debería ser 225, pero es {total_mouse}"
    
    # Teclado: 3*45 = 135
    total_teclado = resultado[resultado['Producto'] == 'Teclado']['Total Ventas'].iloc[0]
    assert total_teclado == 135, f"Teclado debería ser 135, pero es {total_teclado}"
    
    # Verificar que hay 3 productos (Laptop, Mouse, Teclado)
    assert len(resultado) == 3, f"Debería haber 3 productos, pero hay {len(resultado)}"

def test_calcular_totales_con_datos_vacios():
    """Prueba calcular totales con un DataFrame vacío."""
    df_vacio = pd.DataFrame(columns=['Producto', 'Cantidad', 'Precio Unitario'])
    resultado = calcular_totales_por_producto(df_vacio)
    
    assert len(resultado) == 0, "El resultado debería estar vacío"

# ============================================================
# PRUEBAS PARA agregar_total_general()
# ============================================================
def test_agregar_total_general():
    """Prueba que se agregue correctamente la fila de total general."""
    # Crear DataFrame de ejemplo
    df_resumen = pd.DataFrame({
        'Producto': ['Laptop', 'Mouse', 'Teclado'],
        'Total Ventas': [2400, 225, 135]
    })
    
    # Agregar total general
    resultado = agregar_total_general(df_resumen)
    
    # Verificar que se agregó una fila
    assert len(resultado) == 4, f"Debería tener 4 filas, tiene {len(resultado)}"
    
    # Verificar que la última fila es TOTAL GENERAL
    assert resultado['Producto'].iloc[-1] == 'TOTAL GENERAL', "La última fila debería ser TOTAL GENERAL"
    
    # Verificar que el total general es la suma de todos
    assert resultado['Total Ventas'].iloc[-1] == 2760, f"El total general debería ser 2760, es {resultado['Total Ventas'].iloc[-1]}"

def test_agregar_total_general_con_un_producto():
    """Prueba agregar total general con un solo producto."""
    df_resumen = pd.DataFrame({
        'Producto': ['Laptop'],
        'Total Ventas': [2400]
    })
    
    resultado = agregar_total_general(df_resumen)
    
    assert len(resultado) == 2, "Debería tener 2 filas"
    assert resultado['Producto'].iloc[-1] == 'TOTAL GENERAL', "La última fila debería ser TOTAL GENERAL"
    assert resultado['Total Ventas'].iloc[-1] == 2400, "El total general debería ser 2400"

# ============================================================
# PRUEBAS PARA guardar_excel()
# ============================================================
def test_guardar_excel():
    """Prueba guardar un DataFrame en Excel."""
    df = pd.DataFrame({'A': [1, 2], 'B': [3, 4]})
    
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp_path = tmp.name
    
    try:
        # Guardar archivo
        resultado = guardar_excel(df, tmp_path)
        
        # Verificar que se guardó correctamente
        assert resultado is True, "Debería retornar True"
        assert Path(tmp_path).exists(), "El archivo debería existir"
        
        # Verificar que se puede leer el archivo guardado
        df_leido = pd.read_excel(tmp_path)
        assert len(df_leido) == 2, "Debería tener 2 filas"
        assert list(df_leido.columns) == ['A', 'B'], "Columnas incorrectas"
        
    finally:
        if Path(tmp_path).exists():
            os.unlink(tmp_path)

def test_guardar_excel_crea_carpeta():
    """Prueba que guardar_excel cree la carpeta si no existe."""
    # Ruta con carpeta que no existe
    ruta = "carpeta_nueva/archivo.xlsx"
    
    # Crear un DataFrame simple
    df = pd.DataFrame({'A': [1, 2]})
    
    # Guardar (debería crear la carpeta automáticamente)
    resultado = guardar_excel(df, ruta)
    
    # Verificar
    assert resultado is True, "Debería retornar True"
    assert Path(ruta).exists(), "El archivo debería existir"
    
    # Limpiar
    Path(ruta).unlink()
    Path("carpeta_nueva").rmdir()

# ============================================================
# PRUEBAS PARA procesar_excel_ventas()
# ============================================================
def test_procesar_excel_ventas_exitoso():
    """Prueba el proceso completo de generación de reporte."""
    # Crear archivo de entrada temporal
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_in:
        entrada_path = tmp_in.name
        DATOS_PRUEBA.to_excel(entrada_path, index=False)
    
    # Crear ruta de salida temporal
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_out:
        salida_path = tmp_out.name
    
    try:
        # Procesar reporte
        resultado = procesar_excel_ventas(entrada_path, salida_path)
        
        # Verificar resultados
        assert resultado['exito'] is True, "El proceso debería ser exitoso"
        assert resultado['total_general'] == 2760, f"Total general debería ser 2760, es {resultado['total_general']}"
        assert resultado['productos_procesados'] == 3, "Debería tener 3 productos"
        assert resultado['filas_procesadas'] == 5, "Debería haber procesado 5 filas"
        
        # Verificar que el archivo de salida existe
        assert Path(salida_path).exists(), "El archivo de salida debería existir"
        
        # Verificar contenido del archivo de salida
        df_salida = pd.read_excel(salida_path)
        assert len(df_salida) == 4, "Debería tener 4 filas (3 productos + total general)"
        assert df_salida['Producto'].iloc[-1] == 'TOTAL GENERAL', "La última fila debería ser TOTAL GENERAL"
        assert df_salida['Total Ventas'].iloc[-1] == 2760, "El total general debería ser 2760"
        
    finally:
        # Eliminar archivos temporales
        for path in [entrada_path, salida_path]:
            if Path(path).exists():
                os.unlink(path)

def test_procesar_excel_ventas_archivo_inexistente():
    """Prueba que procesar_excel_ventas maneje archivo inexistente."""
    with pytest.raises(FileNotFoundError) as excinfo:
        procesar_excel_ventas("no_existe.xlsx", "salida.xlsx")
    
    assert "no existe" in str(excinfo.value), "Debería indicar que el archivo no existe"

# ============================================================
# EJECUTAR LAS PRUEBAS
# ============================================================
if __name__ == "__main__":
    # Esto permite ejecutar las pruebas directamente con:
    # python tests/test_procesar_excel.py
    pytest.main([__file__, "-v"])
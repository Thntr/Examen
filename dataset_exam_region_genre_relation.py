import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image
import os
import io


def analizar_relacion_region_genero(archivo_excel, archivo_salida="analisis_region_genero.xlsx"):
    """
    Analiza la relación entre región y género de consumo de video

    Args:
        archivo_excel (str): Ruta del archivo Excel de entrada
        archivo_salida (str): Nombre del archivo Excel de salida
    """

    try:
        # Verificar si el archivo existe
        if not os.path.exists(archivo_excel):
            print(f"❌ Error: El archivo '{archivo_excel}' no existe.")
            return None

        # Leer la hoja "Dataset"
        try:
            df = pd.read_excel(archivo_excel, sheet_name='Dataset')
        except:
            # Si no encuentra por nombre, intentar por índice (segunda hoja)
            try:
                df = pd.read_excel(archivo_excel, sheet_name=1)
                print("⚠️  Hoja 'Dataset' no encontrada, usando la segunda hoja del archivo")
            except Exception as e:
                print(f"❌ Error al leer el archivo Excel: {e}")
                return None

        # Verificar si existen las columnas necesarias
        columnas_requeridas = ['REGION', 'GENRE']
        for columna in columnas_requeridas:
            if columna not in df.columns:
                print(f"❌ Error: No se encontró la columna '{columna}' en la hoja")
                print(f"Columnas disponibles: {list(df.columns)}")
                return None

        # Limpiar y preparar los datos
        df_clean = df[['REGION', 'GENRE']].copy()
        df_clean['REGION'] = df_clean['REGION'].astype(str).str.strip()
        df_clean['GENRE'] = df_clean['GENRE'].astype(str).str.strip()
        df_clean = df_clean.dropna()
        df_clean = df_clean[(df_clean['REGION'] != '') & (df_clean['REGION'] != 'nan')]
        df_clean = df_clean[(df_clean['GENRE'] != '') & (df_clean['GENRE'] != 'nan')]

        # Calcular estadísticas generales
        total_registros = len(df_clean)
        regiones_unicas = df_clean['REGION'].nunique()
        generos_unicos = df_clean['GENRE'].nunique()

        print("=" * 70)
        print("🌍 ANÁLISIS DE RELACIÓN REGIÓN-GÉNERO")
        print("=" * 70)
        print(f"📊 Total de registros válidos: {total_registros:,}")
        print(f"🗺️  Regiones únicas: {regiones_unicas}")
        print(f"🎭 Géneros únicos: {generos_unicos}")
        print("=" * 70)

        # Crear tabla de contingencia (frecuencias cruzadas)
        tabla_contingencia = pd.crosstab(df_clean['REGION'], df_clean['GENRE'])

        # Calcular porcentajes por región
        tabla_porcentajes = tabla_contingencia.div(tabla_contingencia.sum(axis=1), axis=0) * 100
        tabla_porcentajes = tabla_porcentajes.round(2)

        # Encontrar top géneros por región
        top_generos_por_region = []

        for region in tabla_contingencia.index:
            generos_region = tabla_contingencia.loc[region]
            total_region = generos_region.sum()

            # Top 3 géneros
            top3 = generos_region.nlargest(3)

            for i, (genero, count) in enumerate(top3.items(), 1):
                porcentaje = (count / total_region) * 100
                top_generos_por_region.append({
                    'REGION': region,
                    'RANKING': i,
                    'GÉNERO': genero,
                    'VISUALIZACIONES': count,
                    'PORCENTAJE': round(porcentaje, 2),
                    'TOTAL_REGION': total_region
                })

        df_top = pd.DataFrame(top_generos_por_region)

        # Estadísticas por región
        stats_region = pd.DataFrame({
            'REGION': tabla_contingencia.index,
            'TOTAL_VISUALIZACIONES': tabla_contingencia.sum(axis=1),
            'GÉNEROS_ÚNICOS': (tabla_contingencia > 0).sum(axis=1),
            'GÉNERO_MÁS_POPULAR': tabla_contingencia.idxmax(axis=1),
            'VISTAS_GÉNERO_TOP': tabla_contingencia.max(axis=1),
            'PORCENTAJE_GÉNERO_TOP': [round((max(row) / sum(row)) * 100, 2) for _, row in tabla_contingencia.iterrows()]
        })

        # Calcular diversidad de consumo (índice de Herfindahl-Hirschman)
        def calcular_hhi(porcentajes):
            return sum([p ** 2 for p in porcentajes])

        stats_region['ÍNDICE_DIVERSIDAD'] = [
            calcular_hhi(tabla_porcentajes.loc[region].values) for region in stats_region['REGION']
        ]

        # Clasificar diversidad
        stats_region['NIVEL_DIVERSIDAD'] = pd.cut(stats_region['ÍNDICE_DIVERSIDAD'],
                                                  bins=[0, 1500, 2500, 10000],
                                                  labels=['Alta Diversidad', 'Media Diversidad', 'Baja Diversidad'])

        # Mostrar resultados en consola
        print("🏆 TOP 1 GÉNERO POR REGIÓN:")
        print("-" * 60)
        top1_regiones = df_top[df_top['RANKING'] == 1].sort_values('PORCENTAJE', ascending=False)

        for i, (_, row) in enumerate(top1_regiones.head(10).iterrows(), 1):
            print(f"{i:2d}. {row['REGION']:<15} → {row['GÉNERO']:<20} ({row['PORCENTAJE']}%)")

        print("=" * 70)
        print("📈 ESTADÍSTICAS GENERALES:")
        print("-" * 40)
        print(f"Región con mayor consumo: {stats_region.loc[stats_region['TOTAL_VISUALIZACIONES'].idxmax(), 'REGION']}")
        print(f"Región con menor consumo: {stats_region.loc[stats_region['TOTAL_VISUALIZACIONES'].idxmin(), 'REGION']}")
        print(f"Género más popular global: {tabla_contingencia.sum().idxmax()}")
        print(f"Región más diversa: {stats_region.loc[stats_region['ÍNDICE_DIVERSIDAD'].idxmin(), 'REGION']}")
        print(f"Región menos diversa: {stats_region.loc[stats_region['ÍNDICE_DIVERSIDAD'].idxmax(), 'REGION']}")

        # Crear archivo Excel con análisis completo
        with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
            # Hoja 1: Tabla de contingencia (frecuencias absolutas)
            tabla_contingencia.to_excel(writer, sheet_name='Frecuencias Absolutas')

            # Hoja 2: Tabla de porcentajes
            tabla_porcentajes.to_excel(writer, sheet_name='Porcentajes por Región')

            # Hoja 3: Top géneros por región
            df_top.to_excel(writer, sheet_name='Top Géneros por Región', index=False)

            # Hoja 4: Estadísticas por región
            stats_region.to_excel(writer, sheet_name='Estadísticas Regiones', index=False)

            # Hoja 5: Resumen ejecutivo
            resumen_data = {
                'MÉTRICA': [
                    'Total de registros analizados',
                    'Regiones únicas identificadas',
                    'Géneros únicos identificados',
                    'Región con mayor consumo total',
                    'Región con menor consumo total',
                    'Género más popular globalmente',
                    'Región con mayor diversidad',
                    'Región con menor diversidad',
                    'Porcentaje promedio del género top',
                    'Regiones con alta diversidad',
                    'Regiones con baja diversidad'
                ],
                'VALOR': [
                    total_registros,
                    regiones_unicas,
                    generos_unicos,
                    stats_region.loc[stats_region['TOTAL_VISUALIZACIONES'].idxmax(), 'REGION'],
                    stats_region.loc[stats_region['TOTAL_VISUALIZACIONES'].idxmin(), 'REGION'],
                    tabla_contingencia.sum().idxmax(),
                    stats_region.loc[stats_region['ÍNDICE_DIVERSIDAD'].idxmin(), 'REGION'],
                    stats_region.loc[stats_region['ÍNDICE_DIVERSIDAD'].idxmax(), 'REGION'],
                    f"{stats_region['PORCENTAJE_GÉNERO_TOP'].mean():.1f}%",
                    len(stats_region[stats_region['NIVEL_DIVERSIDAD'] == 'Alta Diversidad']),
                    len(stats_region[stats_region['NIVEL_DIVERSIDAD'] == 'Baja Diversidad'])
                ]
            }
            df_resumen = pd.DataFrame(resumen_data)
            df_resumen.to_excel(writer, sheet_name='Resumen Ejecutivo', index=False)

            # Crear y guardar gráficos
            if len(stats_region) <= 15:  # Solo si hay pocas regiones
                # Gráfico de heatmap
                plt.figure(figsize=(12, 8))
                import seaborn as sns

                # Normalizar para mejor visualización
                tabla_normalizada = tabla_contingencia.div(tabla_contingencia.sum(axis=1), axis=0)

                sns.heatmap(tabla_normalizada, cmap='YlOrRd', annot=False, fmt='.2f')
                plt.title('Distribución de Géneros por Región (Normalizado)')
                plt.xticks(rotation=45, ha='right')
                plt.tight_layout()

                img_heatmap = io.BytesIO()
                plt.savefig(img_heatmap, format='png', dpi=300, bbox_inches='tight')
                img_heatmap.seek(0)
                plt.close()

            # Ajustar anchos de columnas
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 35)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

        print(f"\n💾 Archivo Excel generado: {archivo_salida}")
        print("📋 Hojas incluidas:")
        print("   - Frecuencias Absolutas: Tabla de contingencia completa")
        print("   - Porcentajes por Región: Distribución porcentual")
        print("   - Top Géneros por Región: Ranking top 3 por región")
        print("   - Estadísticas Regiones: Métricas por región")
        print("   - Resumen Ejecutivo: Hallazgos principales")

        # Respuesta a la pregunta principal
        print("=" * 70)
        print("❓ ¿EXISTE UNA RELACIÓN ENTRE REGIÓN Y GÉNERO?")
        print("=" * 70)

        # Analizar variación entre regiones
        variacion_porcentaje = stats_region['PORCENTAJE_GÉNERO_TOP'].std()

        if variacion_porcentaje > 15:
            print("✅ SÍ, existe una fuerte relación entre región y género")
            print("   📊 Se observan patrones geográficos claros en las preferencias")
        elif variacion_porcentaje > 5:
            print("✅ SÍ, existe una relación moderada entre región y género")
            print("   📊 Hay algunas diferencias regionales en las preferencias")
        else:
            print("❌ NO, no existe una relación fuerte entre región y género")
            print("   📊 Las preferencias son similares across todas las regiones")

        print(f"   📈 Variación en preferencias: {variacion_porcentaje:.1f}%")

        return {
            'tabla_contingencia': tabla_contingencia,
            'tabla_porcentajes': tabla_porcentajes,
            'top_generos': df_top,
            'stats_region': stats_region
        }

    except Exception as e:
        print(f"❌ Error inesperado: {e}")
        import traceback
        traceback.print_exc()
        return None


def main():
    """Función principal"""

    print("🌍 ANALIZADOR DE RELACIÓN REGIÓN-GÉNERO")
    print("=" * 50)

    # Solicitar la ruta del archivo Excel
    archivo = input("Ingrese la ruta del archivo Excel: ").strip().strip('"')

    if not archivo:
        print("❌ Debe ingresar una ruta válida")
        return

    # Nombre del archivo de salida
    archivo_salida = input(
        "Ingrese el nombre del archivo de salida (Enter para 'analisis_region_genero.xlsx'): ").strip()
    if not archivo_salida:
        archivo_salida = "analisis_region_genero.xlsx"

    # Ejecutar el análisis
    resultado = analizar_relacion_region_genero(archivo, archivo_salida)

    if resultado is not None:
        print("\n✅ Análisis completado exitosamente!")


# Ejecutar el script directamente
if __name__ == "__main__":
    main()
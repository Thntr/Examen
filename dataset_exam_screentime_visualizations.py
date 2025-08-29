import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image
import os
import io
from scipy import stats


def analizar_recurrencia_consumo(archivo_excel, archivo_salida="analisis_recurrencia.xlsx"):
    """
    Analiza la recurrencia de consumo por cliente basado en frecuencia y screentime

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
        columnas_requeridas = ['CUSTOMER_ID', 'SCREENTIME']
        for columna in columnas_requeridas:
            if columna not in df.columns:
                print(f"❌ Error: No se encontró la columna '{columna}' en la hoja")
                print(f"Columnas disponibles: {list(df.columns)}")
                return None

        # Limpiar y preparar los datos
        df_clean = df[['CUSTOMER_ID', 'SCREENTIME']].copy()
        df_clean['CUSTOMER_ID'] = df_clean['CUSTOMER_ID'].astype(str).str.strip()

        # Convertir SCREENTIME a numérico, manejando posibles errores
        df_clean['SCREENTIME'] = pd.to_numeric(df_clean['SCREENTIME'], errors='coerce')
        df_clean = df_clean.dropna()

        total_registros = len(df_clean)
        clientes_unicos = df_clean['CUSTOMER_ID'].nunique()

        print("=" * 80)
        print("📊 ANÁLISIS DE RECURRENCIA DE CONSUMO POR CLIENTE")
        print("=" * 80)
        print(f"👥 Total de registros válidos: {total_registros:,}")
        print(f"👤 Clientes únicos: {clientes_unicos}")
        print(f"📈 Ratio promedio de visualizaciones por cliente: {total_registros / clientes_unicos:.2f}")
        print("=" * 80)

        # Análisis de frecuencia de visualizaciones
        frecuencia_clientes = df_clean['CUSTOMER_ID'].value_counts().reset_index()
        frecuencia_clientes.columns = ['CUSTOMER_ID', 'FRECUENCIA_VISUALIZACIONES']

        # Análisis de screentime acumulado
        screentime_clientes = df_clean.groupby('CUSTOMER_ID')['SCREENTIME'].agg([
            'sum', 'mean', 'std', 'count'
        ]).reset_index()
        screentime_clientes.columns = ['CUSTOMER_ID', 'SCREENTIME_TOTAL', 'SCREENTIME_PROMEDIO',
                                       'SCREENTIME_DESVEST', 'FRECUENCIA']

        # Combinar ambos análisis
        analisis_completo = frecuencia_clientes.merge(screentime_clientes, on='CUSTOMER_ID', how='inner')

        # Calcular métricas adicionales
        analisis_completo['SCREENTIME_POR_VISUALIZACION'] = analisis_completo['SCREENTIME_TOTAL'] / analisis_completo[
            'FRECUENCIA_VISUALIZACIONES']

        # Clasificar clientes por frecuencia
        bins_frecuencia = [0, 1, 3, 10, 50, float('inf')]
        labels_frecuencia = ['Ocasional (1)', 'Frecuente (2-3)', 'Muy Frecuente (4-10)',
                             'Super Usuario (11-50)', 'Power User (50+)']

        analisis_completo['CATEGORIA_FRECUENCIA'] = pd.cut(analisis_completo['FRECUENCIA_VISUALIZACIONES'],
                                                           bins=bins_frecuencia,
                                                           labels=labels_frecuencia,
                                                           right=False)

        # Clasificar clientes por screentime total
        percentiles = analisis_completo['SCREENTIME_TOTAL'].quantile([0.33, 0.66])
        bins_screentime = [0, percentiles.iloc[0], percentiles.iloc[1], float('inf')]
        labels_screentime = ['Bajo Consumo', 'Medio Consumo', 'Alto Consumo']

        analisis_completo['CATEGORIA_SCREENTIME'] = pd.cut(analisis_completo['SCREENTIME_TOTAL'],
                                                           bins=bins_screentime,
                                                           labels=labels_screentime)

        # Estadísticas generales
        stats_generales = {
            'Total clientes analizados': len(analisis_completo),
            'Total screentime acumulado': f"{analisis_completo['SCREENTIME_TOTAL'].sum():,.0f} minutos",
            'Screentime promedio por cliente': f"{analisis_completo['SCREENTIME_TOTAL'].mean():.1f} minutos",
            'Visualizaciones promedio por cliente': f"{analisis_completo['FRECUENCIA_VISUALIZACIONES'].mean():.2f}",
            'Screentime promedio por visualización': f"{analisis_completo['SCREENTIME_POR_VISUALIZACION'].mean():.1f} minutos",
            'Cliente con más visualizaciones': f"{analisis_completo.loc[analisis_completo['FRECUENCIA_VISUALIZACIONES'].idxmax(), 'CUSTOMER_ID']} ({analisis_completo['FRECUENCIA_VISUALIZACIONES'].max()} views)",
            'Cliente con mayor screentime': f"{analisis_completo.loc[analisis_completo['SCREENTIME_TOTAL'].idxmax(), 'CUSTOMER_ID']} ({analisis_completo['SCREENTIME_TOTAL'].max():.0f} minutos)"
        }

        # Distribución por categorías
        dist_frecuencia = analisis_completo['CATEGORIA_FRECUENCIA'].value_counts()
        dist_screentime = analisis_completo['CATEGORIA_SCREENTIME'].value_counts()

        # Correlación entre frecuencia y screentime
        correlacion = analisis_completo['FRECUENCIA_VISUALIZACIONES'].corr(analisis_completo['SCREENTIME_TOTAL'])

        print("📈 ESTADÍSTICAS GENERALES:")
        print("-" * 50)
        for key, value in stats_generales.items():
            print(f"• {key}: {value}")

        print("\n🎯 DISTRIBUCIÓN POR FRECUENCIA DE VISUALIZACIONES:")
        print("-" * 60)
        for categoria, count in dist_frecuencia.items():
            porcentaje = (count / len(analisis_completo)) * 100
            print(f"• {categoria}: {count} clientes ({porcentaje:.1f}%)")

        print("\n⏰ DISTRIBUCIÓN POR SCREENTIME TOTAL:")
        print("-" * 50)
        for categoria, count in dist_screentime.items():
            porcentaje = (count / len(analisis_completo)) * 100
            print(f"• {categoria}: {count} clientes ({porcentaje:.1f}%)")

        print(f"\n📊 Correlación frecuencia-screentime: {correlacion:.3f}")

        # Análisis de segmentos combinados
        segmentos = analisis_completo.groupby(['CATEGORIA_FRECUENCIA', 'CATEGORIA_SCREENTIME']).size().unstack(
            fill_value=0)

        print("\n🔍 SEGMENTACIÓN COMBINADA (Frecuencia vs Screentime):")
        print("-" * 60)
        print(segmentos)

        # Identificar clientes valiosos (alta frecuencia + alto screentime)
        clientes_valiosos = analisis_completo[
            (analisis_completo['CATEGORIA_FRECUENCIA'].isin(
                ['Muy Frecuente (4-10)', 'Super Usuario (11-50)', 'Power User (50+)'])) &
            (analisis_completo['CATEGORIA_SCREENTIME'] == 'Alto Consumo')
            ]

        print(f"\n💎 Clientes valiosos (alta frecuencia + alto screentime): {len(clientes_valiosos)}")

        # Crear gráficos
        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(15, 12))

        # Gráfico 1: Distribución de frecuencia
        dist_frecuencia.plot(kind='bar', ax=ax1, color='skyblue', edgecolor='black')
        ax1.set_title('Distribución de Clientes por Frecuencia de Visualizaciones')
        ax1.set_xlabel('Categoría de Frecuencia')
        ax1.set_ylabel('Número de Clientes')
        ax1.tick_params(axis='x', rotation=45)

        # Gráfico 2: Distribución de screentime
        dist_screentime.plot(kind='bar', ax=ax2, color='lightcoral', edgecolor='black')
        ax2.set_title('Distribución de Clientes por Screentime Total')
        ax2.set_xlabel('Categoría de Screentime')
        ax2.set_ylabel('Número de Clientes')
        ax2.tick_params(axis='x', rotation=45)

        # Gráfico 3: Scatter plot frecuencia vs screentime
        ax3.scatter(analisis_completo['FRECUENCIA_VISUALIZACIONES'],
                    analisis_completo['SCREENTIME_TOTAL'],
                    alpha=0.6, color='green', s=30)
        ax3.set_title('Relación entre Frecuencia y Screentime Total')
        ax3.set_xlabel('Frecuencia de Visualizaciones')
        ax3.set_ylabel('Screentime Total (minutos)')
        ax3.set_xscale('log')
        ax3.set_yscale('log')

        # Gráfico 4: Heatmap de segmentación
        im = ax4.imshow(segmentos.values, cmap='YlOrRd', aspect='auto')
        ax4.set_title('Segmentación: Frecuencia vs Screentime')
        ax4.set_xticks(range(len(segmentos.columns)))
        ax4.set_yticks(range(len(segmentos.index)))
        ax4.set_xticklabels(segmentos.columns, rotation=45)
        ax4.set_yticklabels(segmentos.index)
        plt.colorbar(im, ax=ax4)

        plt.tight_layout()

        # Guardar gráficos
        img_buffer = io.BytesIO()
        plt.savefig(img_buffer, format='png', dpi=300, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close()

        # Crear archivo Excel con análisis completo
        with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
            # Hoja 1: Análisis completo por cliente
            analisis_completo.to_excel(writer, sheet_name='Análisis por Cliente', index=False)

            # Hoja 2: Estadísticas generales
            df_stats = pd.DataFrame(list(stats_generales.items()), columns=['Métrica', 'Valor'])
            df_stats.to_excel(writer, sheet_name='Estadísticas', index=False)

            # Hoja 3: Segmentación
            segmentos.reset_index().to_excel(writer, sheet_name='Segmentación', index=False)

            # Hoja 4: Clientes valiosos
            clientes_valiosos.to_excel(writer, sheet_name='Clientes Valiosos', index=False)

            # Hoja 5: Top clientes por diferentes métricas
            top_frecuencia = analisis_completo.nlargest(20, 'FRECUENCIA_VISUALIZACIONES')
            top_screentime = analisis_completo.nlargest(20, 'SCREENTIME_TOTAL')
            top_promedio = analisis_completo.nlargest(20, 'SCREENTIME_PROMEDIO')

            top_frecuencia.to_excel(writer, sheet_name='Top Frecuencia', index=False)
            top_screentime.to_excel(writer, sheet_name='Top Screentime', index=False)
            top_promedio.to_excel(writer, sheet_name='Top Promedio', index=False)

            # Agregar gráficos
            workbook = writer.book
            worksheet = workbook['Estadísticas']

            img = Image(img_buffer)
            img.width = 800
            img.height = 600
            worksheet.add_image(img, 'D2')

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
        print("   - Análisis por Cliente: Datos completos de cada cliente")
        print("   - Estadísticas: Métricas generales + gráficos")
        print("   - Segmentación: Tabla de contingencia")
        print("   - Clientes Valiosos: Clientes más engagados")
        print("   - Top Frecuencia: Clientes con más visualizaciones")
        print("   - Top Screentime: Clientes con mayor tiempo total")
        print("   - Top Promedio: Clientes con mayor tiempo por visualización")

        return analisis_completo

    except Exception as e:
        print(f"❌ Error inesperado: {e}")
        import traceback
        traceback.print_exc()
        return None


def main():
    """Función principal"""

    print("📊 ANALIZADOR DE RECURRENCIA DE CONSUMO")
    print("=" * 50)

    # Solicitar la ruta del archivo Excel
    archivo = input("Ingrese la ruta del archivo Excel: ").strip().strip('"')

    if not archivo:
        print("❌ Debe ingresar una ruta válida")
        return

    # Nombre del archivo de salida
    archivo_salida = input("Ingrese el nombre del archivo de salida (Enter para 'analisis_recurrencia.xlsx'): ").strip()
    if not archivo_salida:
        archivo_salida = "analisis_recurrencia.xlsx"

    # Ejecutar el análisis
    resultado = analizar_recurrencia_consumo(archivo, archivo_salida)

    if resultado is not None:
        print("\n✅ Análisis completado exitosamente!")


# Ejecutar el script directamente
if __name__ == "__main__":
    main()
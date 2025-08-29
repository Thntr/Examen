import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image
import os
import io


def analizar_shows_por_visualizaciones(archivo_excel, archivo_salida="analisis_shows_tv.xlsx", top_n=20):
    """
    Analiza los shows de TV únicos ordenados por visualizaciones, incluyendo su género

    Args:
        archivo_excel (str): Ruta del archivo Excel de entrada
        archivo_salida (str): Nombre del archivo Excel de salida
        top_n (int): Número de shows top a incluir en el gráfico
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
        columnas_requeridas = ['TITLE', 'GENRE']
        for columna in columnas_requeridas:
            if columna not in df.columns:
                print(f"❌ Error: No se encontró la columna '{columna}' en la hoja")
                print(f"Columnas disponibles: {list(df.columns)}")
                return None

        # Limpiar y preparar los datos
        df_clean = df[['TITLE', 'GENRE']].copy()
        df_clean['TITLE'] = df_clean['TITLE'].astype(str).str.strip()
        df_clean['GENRE'] = df_clean['GENRE'].astype(str).str.strip()
        df_clean = df_clean.dropna()
        df_clean = df_clean[(df_clean['TITLE'] != '') & (df_clean['TITLE'] != 'nan')]
        df_clean = df_clean[(df_clean['GENRE'] != '') & (df_clean['GENRE'] != 'nan')]

        # Calcular estadísticas generales
        total_visualizaciones = len(df_clean)
        shows_unicos = df_clean['TITLE'].nunique()
        generos_unicos = df_clean['GENRE'].nunique()

        print("=" * 70)
        print("📺 ANÁLISIS DE SHOWS DE TV POR VISUALIZACIONES")
        print("=" * 70)
        print(f"📊 Total de visualizaciones: {total_visualizaciones:,}")
        print(f"🎬 Shows únicos: {shows_unicos}")
        print(f"🎭 Géneros únicos: {generos_unicos}")
        print("=" * 70)

        # Contar visualizaciones por show y género
        conteo_shows = df_clean.groupby(['TITLE', 'GENRE']).size().reset_index(name='VISUALIZACIONES')

        # Ordenar por visualizaciones (descendente)
        conteo_shows = conteo_shows.sort_values('VISUALIZACIONES', ascending=False)

        # Agregar porcentaje
        conteo_shows['PORCENTAJE'] = (conteo_shows['VISUALIZACIONES'] / total_visualizaciones * 100).round(4)
        conteo_shows['PORCENTAJE_ACUMULADO'] = conteo_shows['PORCENTAJE'].cumsum()

        # Calcular estadísticas adicionales
        shows_top_10 = conteo_shows.head(10)
        porcentaje_top_10 = shows_top_10['PORCENTAJE'].sum()

        # Encontrar el show más popular de cada género
        top_show_por_genero = conteo_shows.loc[conteo_shows.groupby('GENRE')['VISUALIZACIONES'].idxmax()]

        print(f"🏆 TOP 10 SHOWS MÁS VISTOS ({porcentaje_top_10:.1f}% del total):")
        print("-" * 80)
        for i, (_, row) in enumerate(shows_top_10.iterrows(), 1):
            print(
                f"{i:2d}. {row['TITLE'][:40]:<40} [{row['GENRE']:<12}] {row['VISUALIZACIONES']:>6,} views ({row['PORCENTAJE']:.2f}%)")

        print("=" * 70)
        print("🎯 SHOW MÁS POPULAR POR GÉNERO:")
        print("-" * 60)
        for _, row in top_show_por_genero.iterrows():
            print(f"• {row['GENRE']:<15}: {row['TITLE'][:30]:<30} ({row['VISUALIZACIONES']:,} views)")

        # Preparar datos para el gráfico de pastel
        if len(conteo_shows) > top_n:
            # Agrupar shows menos populares en "Otros"
            top_shows = conteo_shows.head(top_n)
            otros_visualizaciones = conteo_shows['VISUALIZACIONES'].iloc[top_n:].sum()
            otros_porcentaje = conteo_shows['PORCENTAJE'].iloc[top_n:].sum()

            datos_grafico = pd.concat([
                top_shows[['TITLE', 'VISUALIZACIONES', 'PORCENTAJE']],
                pd.DataFrame({
                    'TITLE': ['Otros'],
                    'VISUALIZACIONES': [otros_visualizaciones],
                    'PORCENTAJE': [otros_porcentaje]
                })
            ])
        else:
            datos_grafico = conteo_shows.copy()

        # Crear etiquetas para el gráfico
        etiquetas = []
        for _, row in datos_grafico.iterrows():
            if row['TITLE'] == 'Otros':
                etiquetas.append(f"Otros ({row['PORCENTAJE']:.1f}%)")
            else:
                etiqueta_corta = row['TITLE'][:15] + '...' if len(row['TITLE']) > 15 else row['TITLE']
                etiquetas.append(f"{etiqueta_corta} ({row['PORCENTAJE']:.1f}%)")

        # Crear gráfico de pastel
        plt.figure(figsize=(14, 10))

        colors = plt.cm.tab20c(range(len(datos_grafico)))
        wedges, texts, autotexts = plt.pie(
            datos_grafico['VISUALIZACIONES'],
            labels=etiquetas,
            autopct='%1.1f%%',
            startangle=90,
            colors=colors,
            pctdistance=0.85,
            textprops={'fontsize': 8}
        )

        # Mejorar la legibilidad
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
            autotext.set_fontsize(7)

        plt.title(f'Distribución de Visualizaciones por Show de TV\n(Top {top_n} + Otros)', fontsize=16,
                  fontweight='bold', pad=20)
        plt.axis('equal')

        # Agregar leyenda con información completa
        legend_labels = []
        for _, row in datos_grafico.iterrows():
            if row['TITLE'] == 'Otros':
                legend_labels.append(f"Otros: {row['VISUALIZACIONES']:,} views ({row['PORCENTAJE']:.2f}%)")
            else:
                # Encontrar el género del show
                genero = conteo_shows[conteo_shows['TITLE'] == row['TITLE']]['GENRE'].iloc[0]
                legend_labels.append(f"{row['TITLE']} [{genero}]: {row['VISUALIZACIONES']:,} views")

        plt.legend(wedges, legend_labels, title="Shows Detallados", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1),
                   fontsize=8)

        # Guardar el gráfico en memoria
        img_buffer = io.BytesIO()
        plt.savefig(img_buffer, format='png', dpi=300, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close()

        # Crear archivo Excel con análisis completo
        with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
            # Hoja 1: Shows ordenados por visualizaciones
            conteo_shows.to_excel(writer, sheet_name='Shows por Visualizaciones', index=False)

            # Hoja 2: Top shows por género
            top_show_por_genero.to_excel(writer, sheet_name='Top por Género', index=False)

            # Hoja 3: Estadísticas generales
            stats_data = {
                'ESTADÍSTICA': [
                    'Total de visualizaciones',
                    'Shows de TV únicos',
                    'Géneros únicos',
                    'Show más visto',
                    'Visualizaciones del show top',
                    'Porcentaje del show top',
                    'Género más popular',
                    'Show más visto del género más popular',
                    'Porcentaje acumulado top 10 shows',
                    'Porcentaje acumulado top 20 shows',
                    'Shows con solo 1 visualización'
                ],
                'VALOR': [
                    total_visualizaciones,
                    shows_unicos,
                    generos_unicos,
                    conteo_shows.iloc[0]['TITLE'],
                    conteo_shows.iloc[0]['VISUALIZACIONES'],
                    f"{conteo_shows.iloc[0]['PORCENTAJE']:.4f}%",
                    conteo_shows.groupby('GENRE')['VISUALIZACIONES'].sum().idxmax(),
                    top_show_por_genero[top_show_por_genero['GENRE'] ==
                                        conteo_shows.groupby('GENRE')['VISUALIZACIONES'].sum().idxmax()]['TITLE'].iloc[
                        0],
                    f"{conteo_shows.head(10)['PORCENTAJE'].sum():.2f}%",
                    f"{conteo_shows.head(20)['PORCENTAJE'].sum():.2f}%",
                    len(conteo_shows[conteo_shows['VISUALIZACIONES'] == 1])
                ]
            }
            df_stats = pd.DataFrame(stats_data)
            df_stats.to_excel(writer, sheet_name='Estadísticas', index=False)

            # Hoja 4: Distribución por género (para contexto)
            distribucion_genero = df_clean['GENRE'].value_counts().reset_index()
            distribucion_genero.columns = ['GÉNERO', 'VISUALIZACIONES']
            distribucion_genero['PORCENTAJE'] = (
                        distribucion_genero['VISUALIZACIONES'] / total_visualizaciones * 100).round(2)
            distribucion_genero.to_excel(writer, sheet_name='Distribución por Género', index=False)

            # Obtener el workbook para agregar el gráfico
            workbook = writer.book
            worksheet = workbook['Shows por Visualizaciones']

            # Insertar el gráfico en el Excel
            img = Image(img_buffer)
            img.width = 800
            img.height = 600

            # Agregar imagen después de los datos
            max_row = len(conteo_shows) + 4
            worksheet.add_image(img, f'F{max_row}')

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
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

        print(f"\n💾 Archivo Excel generado: {archivo_salida}")
        print("📋 Hojas incluidas:")
        print("   - Shows por Visualizaciones: Lista completa ordenada + gráfico")
        print("   - Top por Género: Show más popular de cada género")
        print("   - Estadísticas: Métricas clave del análisis")
        print("   - Distribución por Género: Contexto general de géneros")

        return conteo_shows

    except Exception as e:
        print(f"❌ Error inesperado: {e}")
        import traceback
        traceback.print_exc()
        return None


def main():
    """Función principal"""

    print("📺 ANALIZADOR DE SHOWS DE TV POR VISUALIZACIONES")
    print("=" * 50)

    # Solicitar la ruta del archivo Excel
    archivo = input("Ingrese la ruta del archivo Excel: ").strip().strip('"')

    if not archivo:
        print("❌ Debe ingresar una ruta válida")
        return

    # Nombre del archivo de salida
    archivo_salida = input("Ingrese el nombre del archivo de salida (Enter para 'analisis_shows_tv.xlsx'): ").strip()
    if not archivo_salida:
        archivo_salida = "analisis_shows_tv.xlsx"

    # Número de shows para el top
    try:
        top_n = int(input("¿Cuántos shows top incluir en el gráfico? (Enter para 20): ") or "20")
    except:
        top_n = 20
        print("Usando valor por defecto: 20")

    # Ejecutar el análisis
    resultado = analizar_shows_por_visualizaciones(archivo, archivo_salida, top_n)

    if resultado is not None:
        print("\n✅ Análisis completado exitosamente!")


# Ejecutar el script directamente
if __name__ == "__main__":
    main()
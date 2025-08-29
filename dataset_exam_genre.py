import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils.dataframe import dataframe_to_rows
import os
import io


def analizar_generos_y_grafico(archivo_excel, archivo_salida="analisis_generos.xlsx"):
    """
    Analiza los géneros más vistos y genera un gráfico de pastel en Excel

    Args:
        archivo_excel (str): Ruta del archivo Excel de entrada
        archivo_salida (str): Nombre del archivo Excel de salida
    """

    try:
        # Verificar si el archivo existe
        if not os.path.exists(archivo_excel):
            print(f"❌ Error: El archivo '{archivo_excel}' no existe.")
            return

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
                return

        # Verificar si existe la columna GENRE
        if 'GENRE' not in df.columns:
            print("❌ Error: No se encontró la columna 'GENRE' en la hoja")
            print(f"Columnas disponibles: {list(df.columns)}")
            return

        # Obtener y limpiar los datos de género
        generos = df['GENRE'].dropna().astype(str).str.strip()

        # Contar frecuencias
        conteo_generos = generos.value_counts()
        total_registros = len(generos)

        print("=" * 60)
        print("🎬 ANÁLISIS DE GÉNEROS DE VIDEOS")
        print("=" * 60)
        print(f"📊 Total de visualizaciones: {total_registros}")
        print(f"🎭 Géneros únicos encontrados: {len(conteo_generos)}")
        print("=" * 60)

        # Mostrar el top 10 de géneros
        print("🏆 TOP 10 GÉNEROS MÁS VISTOS:")
        print("-" * 40)
        for i, (genero, count) in enumerate(conteo_generos.head(10).items(), 1):
            porcentaje = (count / total_registros) * 100
            print(f"{i:2d}. {genero:<25} {count:>5} views ({porcentaje:.1f}%)")

        # Género más visto
        genero_mas_visto = conteo_generos.index[0]
        vistas_genero_mas_visto = conteo_generos.iloc[0]
        porcentaje_mas_visto = (vistas_genero_mas_visto / total_registros) * 100

        print("=" * 60)
        print(f"👑 GÉNERO MÁS VISTO: '{genero_mas_visto}'")
        print(f"   👀 Visualizaciones: {vistas_genero_mas_visto}")
        print(f"   📈 Porcentaje: {porcentaje_mas_visto:.2f}%")
        print("=" * 60)

        # Crear DataFrame para el reporte
        reporte_df = pd.DataFrame({
            'GÉNERO': conteo_generos.index,
            'VISUALIZACIONES': conteo_generos.values,
            'PORCENTAJE (%)': (conteo_generos.values / total_registros * 100).round(2)
        })

        # Crear gráfico de pastel
        plt.figure(figsize=(12, 8))

        # Si hay muchos géneros, agrupar los menos frecuentes en "Otros"
        if len(conteo_generos) > 10:
            top_10 = conteo_generos.head(10)
            otros = conteo_generos[10:].sum()
            datos_grafico = pd.concat([top_10, pd.Series([otros], index=['Otros'])])
            etiquetas = list(top_10.index) + ['Otros']
        else:
            datos_grafico = conteo_generos
            etiquetas = conteo_generos.index

        # Crear el gráfico de pastel
        colors = plt.cm.Set3(range(len(datos_grafico)))
        wedges, texts, autotexts = plt.pie(datos_grafico.values,
                                           labels=etiquetas,
                                           autopct='%1.1f%%',
                                           startangle=90,
                                           colors=colors,
                                           textprops={'fontsize': 9})

        plt.title('Distribución de Visualizaciones por Género\n', fontsize=16, fontweight='bold')
        plt.axis('equal')  # Asegura que el pastel sea circular

        # Mejorar la legibilidad
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')

        # Guardar el gráfico en memoria
        img_buffer = io.BytesIO()
        plt.savefig(img_buffer, format='png', dpi=300, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close()

        # Crear archivo Excel de salida
        with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
            # Guardar el reporte completo
            reporte_df.to_excel(writer, sheet_name='Reporte Géneros', index=False)

            # Guardar estadísticas resumidas
            resumen_df = pd.DataFrame({
                'ESTADÍSTICA': ['Total visualizaciones', 'Géneros únicos', 'Género más visto',
                                'Vistas género más visto', 'Porcentaje género más visto'],
                'VALOR': [total_registros, len(conteo_generos), genero_mas_visto,
                          vistas_genero_mas_visto, f'{porcentaje_mas_visto:.2f}%']
            })
            resumen_df.to_excel(writer, sheet_name='Resumen', index=False)

            # Obtener el workbook y la hoja
            workbook = writer.book
            worksheet = workbook['Reporte Géneros']

            # Insertar el gráfico en el Excel
            img = Image(img_buffer)
            img.width = 600
            img.height = 400

            # Agregar imagen después de los datos
            max_row = len(reporte_df) + 2
            worksheet.add_image(img, f'E{max_row}')

            # Ajustar el ancho de las columnas
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

        print(f"💾 Archivo Excel generado: {archivo_salida}")
        print("📊 Gráfico de pastel incluido en el archivo")

        return reporte_df

    except Exception as e:
        print(f"❌ Error inesperado: {e}")
        return None


def main():
    """Función principal"""

    print("🎬 ANALIZADOR DE GÉNEROS DE VIDEOS")
    print("=" * 50)

    # Solicitar la ruta del archivo Excel
    archivo = input("Ingrese la ruta del archivo Excel: ").strip().strip('"')

    if not archivo:
        print("❌ Debe ingresar una ruta válida")
        return

    # Nombre del archivo de salida
    archivo_salida = input("Ingrese el nombre del archivo de salida (Enter para 'analisis_generos.xlsx'): ").strip()
    if not archivo_salida:
        archivo_salida = "analisis_generos.xlsx"

    # Ejecutar el análisis
    resultado = analizar_generos_y_grafico(archivo, archivo_salida)

    if resultado is not None:
        print("\n✅ Análisis completado exitosamente!")


# Ejecutar el script directamente
if __name__ == "__main__":
    main()
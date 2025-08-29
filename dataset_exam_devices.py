import pandas as pd
from collections import defaultdict
import os


def analizar_dispositivos_por_cliente(archivo_excel, archivo_salida="analisis_dispositivos.xlsx"):
    """
    Analiza si los clientes utilizan múltiples dispositivos para consumir video

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
        columnas_requeridas = ['CUSTOMER_ID', 'DEVICE']
        for columna in columnas_requeridas:
            if columna not in df.columns:
                print(f"❌ Error: No se encontró la columna '{columna}' en la hoja")
                print(f"Columnas disponibles: {list(df.columns)}")
                return None

        # Limpiar y preparar los datos
        df_clean = df[['CUSTOMER_ID', 'DEVICE']].copy()
        df_clean['CUSTOMER_ID'] = df_clean['CUSTOMER_ID'].astype(str).str.strip()
        df_clean['DEVICE'] = df_clean['DEVICE'].astype(str).str.strip()
        df_clean = df_clean.dropna()

        # Crear diccionario para almacenar dispositivos por cliente
        clientes_dispositivos = defaultdict(set)
        clientes_registros = defaultdict(list)

        # Procesar cada registro
        for _, row in df_clean.iterrows():
            customer_id = row['CUSTOMER_ID']
            device = row['DEVICE']

            clientes_dispositivos[customer_id].add(device)
            clientes_registros[customer_id].append(device)

        # Crear DataFrame con el análisis
        analisis_data = []

        for customer_id, dispositivos in clientes_dispositivos.items():
            cantidad_dispositivos = len(dispositivos)
            lista_dispositivos = ', '.join(sorted(dispositivos))
            total_registros = len(clientes_registros[customer_id])

            analisis_data.append({
                'CUSTOMER_ID': customer_id,
                'CANTIDAD_DISPOSITIVOS': cantidad_dispositivos,
                'DISPOSITIVOS_UTILIZADOS': lista_dispositivos,
                'USA_MULTIPLES_DISPOSITIVOS': 'Sí' if cantidad_dispositivos > 1 else 'No',
                'TOTAL_VISUALIZACIONES': total_registros,
                'DISP_UNICOS/TOTAL': f"{cantidad_dispositivos}/{total_registros}"
            })

        # Crear DataFrame principal
        df_analisis = pd.DataFrame(analisis_data)

        # Ordenar por cantidad de dispositivos (descendente)
        df_analisis = df_analisis.sort_values('CANTIDAD_DISPOSITIVOS', ascending=False)

        # Estadísticas generales
        total_clientes = len(df_analisis)
        clientes_multiples = len(df_analisis[df_analisis['CANTIDAD_DISPOSITIVOS'] > 1])
        porcentaje_multiples = (clientes_multiples / total_clientes * 100) if total_clientes > 0 else 0

        clientes_un_dispositivo = len(df_analisis[df_analisis['CANTIDAD_DISPOSITIVOS'] == 1])
        clientes_mas_de_dos = len(df_analisis[df_analisis['CANTIDAD_DISPOSITIVOS'] > 2])

        # Estadísticas de dispositivos
        todos_dispositivos = [device for devices in clientes_dispositivos.values() for device in devices]
        dispositivo_counts = pd.Series(todos_dispositivos).value_counts()

        print("=" * 70)
        print("📱 ANÁLISIS DE DISPOSITIVOS POR CLIENTE")
        print("=" * 70)
        print(f"👥 Total de clientes únicos: {total_clientes}")
        print(f"📊 Clientes con múltiples dispositivos: {clientes_multiples} ({porcentaje_multiples:.1f}%)")
        print(f"📱 Clientes con un solo dispositivo: {clientes_un_dispositivo}")
        print(f"🔥 Clientes con más de 2 dispositivos: {clientes_mas_de_dos}")
        print("=" * 70)

        if clientes_multiples > 0:
            print("🏆 TOP 10 CLIENTES CON MÁS DISPOSITIVOS:")
            print("-" * 60)
            for i, (_, row) in enumerate(df_analisis.head(10).iterrows(), 1):
                print(
                    f"{i:2d}. {row['CUSTOMER_ID']:<15} {row['CANTIDAD_DISPOSITIVOS']:>2} dispositivos: {row['DISPOSITIVOS_UTILIZADOS']}")

        print("=" * 70)
        print("📊 DISTRIBUCIÓN DE DISPOSITIVOS:")
        print("-" * 40)
        for dispositivo, count in dispositivo_counts.head(10).items():
            print(f"• {dispositivo:<20}: {count:>4} clientes")

        if len(dispositivo_counts) > 10:
            print(f"  ... y {len(dispositivo_counts) - 10} dispositivos más")

        # Respuesta a la pregunta principal
        print("=" * 70)
        print("❓ ¿UTILIZAN MÁS DE UN DISPOSITIVO PARA CONSUMIR VIDEO?")
        print("=" * 70)

        if clientes_multiples > 0:
            print(f"✅ SÍ, {clientes_multiples} clientes ({porcentaje_multiples:.1f}%) utilizan múltiples dispositivos")
            print(f"   Esto representa una parte significativa de la base de clientes")
        else:
            print("❌ NO, ningún cliente utiliza múltiples dispositivos")
            print("   Todos los clientes usan un único dispositivo")

        # Crear archivo Excel con análisis detallado
        with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
            # Hoja principal con análisis completo
            df_analisis.to_excel(writer, sheet_name='Análisis por Cliente', index=False)

            # Hoja con estadísticas resumidas
            stats_data = {
                'ESTADÍSTICA': [
                    'Total clientes únicos',
                    'Clientes con múltiples dispositivos',
                    'Porcentaje con múltiples dispositivos',
                    'Clientes con un solo dispositivo',
                    'Clientes con más de 2 dispositivos',
                    'Dispositivos únicos encontrados'
                ],
                'VALOR': [
                    total_clientes,
                    clientes_multiples,
                    f'{porcentaje_multiples:.1f}%',
                    clientes_un_dispositivo,
                    clientes_mas_de_dos,
                    len(dispositivo_counts)
                ]
            }
            df_stats = pd.DataFrame(stats_data)
            df_stats.to_excel(writer, sheet_name='Estadísticas', index=False)

            # Hoja con top clientes con múltiples dispositivos
            df_multiples = df_analisis[df_analisis['CANTIDAD_DISPOSITIVOS'] > 1].copy()
            df_multiples = df_multiples.sort_values('CANTIDAD_DISPOSITIVOS', ascending=False)
            df_multiples.to_excel(writer, sheet_name='Clientes Múltiples', index=False)

            # Hoja con distribución de dispositivos
            df_dispositivos = pd.DataFrame({
                'DISPOSITIVO': dispositivo_counts.index,
                'CLIENTES_QUE_USAN': dispositivo_counts.values,
                'PORCENTAJE': (dispositivo_counts.values / total_clientes * 100).round(2)
            })
            df_dispositivos.to_excel(writer, sheet_name='Dispositivos', index=False)

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
        print("   - Análisis por Cliente: Detalle completo de cada cliente")
        print("   - Estadísticas: Resumen general")
        print("   - Clientes Múltiples: Solo clientes con >1 dispositivo")
        print("   - Dispositivos: Distribución de dispositivos")

        return df_analisis

    except Exception as e:
        print(f"❌ Error inesperado: {e}")
        import traceback
        traceback.print_exc()
        return None


def main():
    """Función principal"""

    print("📱 ANALIZADOR DE DISPOSITIVOS POR CLIENTE")
    print("=" * 50)

    # Solicitar la ruta del archivo Excel
    archivo = input("Ingrese la ruta del archivo Excel: ").strip().strip('"')

    if not archivo:
        print("❌ Debe ingresar una ruta válida")
        return

    # Nombre del archivo de salida
    archivo_salida = input(
        "Ingrese el nombre del archivo de salida (Enter para 'analisis_dispositivos.xlsx'): ").strip()
    if not archivo_salida:
        archivo_salida = "analisis_dispositivos.xlsx"

    # Ejecutar el análisis
    resultado = analizar_dispositivos_por_cliente(archivo, archivo_salida)

    if resultado is not None:
        print("\n✅ Análisis completado exitosamente!")


# Ejecutar el script directamente
if __name__ == "__main__":
    main()
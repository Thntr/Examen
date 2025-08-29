import pandas as pd
from collections import Counter
import os


def analizar_customer_ids(archivo_excel):
    """
    Analiza los CUSTOMER_ID en un archivo Excel

    Args:
        archivo_excel (str): Ruta del archivo Excel a analizar
    """

    try:
        # Verificar si el archivo existe
        if not os.path.exists(archivo_excel):
            print(f"❌ Error: El archivo '{archivo_excel}' no existe.")
            return

        # Leer la segunda hoja (índice 1) llamada "Dataset"
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

        # Verificar si existe la columna CUSTOMER_ID
        if 'CUSTOMER_ID' not in df.columns:
            print("❌ Error: No se encontró la columna 'CUSTOMER_ID' en la hoja")
            print(f"Columnas disponibles: {list(df.columns)}")
            return

        # Obtener todos los CUSTOMER_ID (eliminando valores nulos)
        customer_ids = df['CUSTOMER_ID'].dropna().astype(str)

        # Contar total de registros y valores únicos
        total_registros = len(customer_ids)
        total_unicos = customer_ids.nunique()

        print("=" * 50)
        print("📊 ANÁLISIS DE CUSTOMER_ID")
        print("=" * 50)
        print(f"📋 Total de registros: {total_registros}")
        print(f"👤 Customer IDs únicos: {total_unicos}")
        print(f"🔄 Registros duplicados: {total_registros - total_unicos}")
        print("=" * 50)

        # Encontrar duplicados y contar repeticiones
        contador = Counter(customer_ids)
        duplicados = {k: v for k, v in contador.items() if v > 1}

        if duplicados:
            print("🔍 CUSTOMER_IDs DUPLICADOS:")
            print("=" * 50)

            # Ordenar por cantidad de repeticiones (descendente)
            duplicados_ordenados = sorted(duplicados.items(), key=lambda x: x[1], reverse=True)

            for customer_id, count in duplicados_ordenados:
                print(f"• {customer_id}: se repite {count} veces")

            print("=" * 50)
            print(f"📈 Total de IDs duplicados: {len(duplicados)}")
        else:
            print("✅ No se encontraron CUSTOMER_IDs duplicados")

        # Mostrar algunos ejemplos únicos (si hay muchos)
        if total_unicos > 0:
            print(f"\n🔹 Primeros 5 CUSTOMER_IDs únicos (ejemplo):")
            unicos = customer_ids.unique()
            for i, uid in enumerate(unicos[:5]):
                print(f"  {i + 1}. {uid}")

            if total_unicos > 5:
                print(f"  ... y {total_unicos - 5} más")

        # Generar resumen estadístico
        print("\n" + "=" * 50)
        print("📈 RESUMEN ESTADÍSTICO")
        print("=" * 50)
        print(f"Porcentaje de duplicados: {(len(duplicados) / total_unicos * 100):.2f}%")
        print(f"Ratio duplicados/únicos: {len(duplicados)}/{total_unicos}")

        # Guardar resultados en un archivo CSV
        try:
            # Crear DataFrame con los resultados
            resultados = pd.DataFrame({
                'CUSTOMER_ID': list(contador.keys()),
                'FRECUENCIA': list(contador.values()),
                'ES_DUPLICADO': [v > 1 for v in contador.values()]
            })

            # Guardar en CSV
            nombre_archivo_salida = 'analisis_customer_ids.csv'
            resultados.to_csv(nombre_archivo_salida, index=False, encoding='utf-8')
            print(f"\n💾 Resultados guardados en: {nombre_archivo_salida}")

        except Exception as e:
            print(f"⚠️  No se pudo guardar el archivo CSV: {e}")

    except Exception as e:
        print(f"❌ Error inesperado: {e}")


def main():
    """Función principal"""

    print("🔍 ANALIZADOR DE CUSTOMER_ID EN EXCEL")
    print("=" * 50)

    # Solicitar la ruta del archivo Excel
    archivo = input("Ingrese la ruta del archivo Excel: ").strip().strip('"')

    if not archivo:
        print("❌ Debe ingresar una ruta válida")
        return

    # Ejecutar el análisis
    analizar_customer_ids(archivo)


# Ejecutar el script directamente
if __name__ == "__main__":
    main()

# Examen
Scripts relativos a una postulación llamada 'Exámen.xlsx'

# Scripts de Análisis para Postulación

## Aviso de Privacidad 📌
Este repositorio **no contiene datos sensibles ni confidenciales**.  
Por cuestiones de privacidad, **no se incluye el dataset original** proporcionado por la empresa (denominado *"Exámen.xlsx"*).  

En este repositorio únicamente se encuentran los **scripts en Python** utilizados para:
- Manipular y procesar cualquier dataset estructurado en formato `.xlsx` o `.csv`.
- Reproducir el flujo de trabajo aplicado durante la postulación.
- Mostrar de manera clara y transparente el proceso de análisis realizado.

⚠️ **Importante**: El nombre de la empresa y los datos originales se mantienen en confidencialidad. Los scripts aquí presentados funcionan de manera general sobre cualquier dataset con características similares.

---

## Propósito 🎯
El objetivo de este repositorio es **brindar al equipo de reclutamiento** una visión detallada de cómo fue realizado el análisis de su dataset de postulación, garantizando transparencia en la metodología empleada.

---

## Contenido del Repositorio 📂
- `dataset_exam_customers.py` → Script empleado para generar los excel con la segmentación de datos, para la primera pregunta: ¿Cuántos clientes consumen video en el mes?
- `dataset_exam_devices.py` → Script empleado para determinar el consumo por tipo de dispositivo.
- `dataset_exam_genre.py` → Script empleado para: ¿Cuál fue el género más visto?
- `dataset_exam_region_genre_relation.py` → Script que determina las relaciones entre género y región.
- `dataset_exam_screentime_visualizations.py` → Este, detecta a los clientes más valiosos según su frecuencia y tiempo en pantalla.
- 'dataset_exam_top_shows.py' → Aquí esta la herramienta que nos permite realizar un top de los Shows más vistos.


---

## Requisitos ⚙️
- Python 3.8+
- Bibliotecas principales:
  - `pandas`
  - `numpy`
  - `matplotlib`
  - `seaborn`
  - `openpyxl`

Instalación rápida de dependencias:
```bash
pip install -r requirements.txt


# 🏛️ Automatizador y Actualizador de Ordenanzas Fiscales

Script desarrollado en **Python** para la lectura, cálculo y actualización masiva de cuadros tarifarios y normativas municipales directamente sobre documentos de Microsoft Word (`.docx`).

## 📋 Descripción del Proyecto
Este proyecto resuelve el complejo desafío operativo de actualizar valores monetarios dispersos en documentos legales extensos. El script escanea documentos DOCX en busca de montos financieros, identifica aquellos marcados mediante resaltado, aplica un porcentaje de aumento ingresado por el usuario y normaliza el formato de salida. 

Al automatizar este proceso, se elimina la carga manual, se reduce a cero el margen de error humano en los cálculos y se preserva intacta la estructura y el formato legal del documento original.

## ⚙️ Características Principales
- **Procesamiento Selectivo:** Escanea párrafos y tablas, modificando de forma segura únicamente los valores numéricos que el usuario haya resaltado en el documento original.
- **Detección Inteligente (Regex):** Utiliza expresiones regulares para capturar y normalizar una amplia variedad de formatos monetarios sucios o inconsistentes (ej: `$ 2312.21`, `$2.732,00`, `44,31`).
- **Formateo Estricto:** Recalcula los montos aplicando redondeo y los reescribe utilizando el estándar de formato local (ej: `1.234,00`).
- **Auditoría y Trazabilidad:** Genera automáticamente un archivo `.csv` de auditoría detallando la ubicación, el valor original, el cálculo exacto y el nuevo valor insertado, además de un log de ejecución en `.txt`.

## 🛠️ Stack Tecnológico
- **Lenguaje:** Python 3.x
- **Librerías principales:** `python-docx` (manipulación de documentos), `re` (expresiones regulares), `csv`, `math`, `pathlib`.

## 🚀 Cómo probar el proyecto localmente

1. Clonar el repositorio:
   ```bash
   git clone [https://github.com/riquelmematias7/etl-ordenanzas-merge.git](https://github.com/riquelmematias7/etl-ordenanzas-merge.git)

2. Instalar las dependencias necesarias:

pip install python-docx

3. Colocar los archivos .docx a procesar en el directorio raíz. Asegurarse de que los valores a actualizar estén resaltados.

4. Ejecutar el script principal:
   
python ordenanzas_merge.py

5. Ingresar el porcentaje de aumento cuando el sistema lo solicite (ej: 8.61).
6. Revisar el nuevo documento generado y el archivo de auditoría CSV en la misma carpeta.

Desarrollado con fines de demostración técnica para portfolio profesional.

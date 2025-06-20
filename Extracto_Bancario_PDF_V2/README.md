# Macro Excel - Conciliación Bancaria desde Extractos PDF (Versión Alternativa)

Esta macro en VBA para Excel automatiza el proceso de conciliación de extractos bancarios importados desde archivos PDF.

## 📌 Contexto y propósito

Esta solución fue desarrollada para dar soporte al tratamiento automatizado de extractos bancarios con estructuras diferentes entre empresas.  
En este caso, el archivo se deriva de un **PDF bancario importado a Excel**, y por razones de confidencialidad, **los datos reales no están incluidos en este repositorio**.

---

## ✅ ¿Qué hace esta macro?

- Detecta dinámicamente los valores en la columna **E**, que representa los movimientos del extracto.
- Separa los valores **positivos** en la columna **H** y los **negativos convertidos a positivos** en la columna **I**.
- Compara los **positivos** con los valores en la columna **A** (débitos).
- Compara los **negativos (convertidos)** con los valores en la columna **B** (créditos).
- Genera dos hojas automáticas con los valores que no tienen coincidencia:
  - `NoEnDébitos` → valores positivos sin encontrar en A.
  - `NoEnCréditos` → valores negativos (convertidos) sin encontrar en B.
- Todo el proceso se adapta automáticamente a la longitud real de los datos (no requiere modificar la macro al cambiar el volumen).

---

## 📂 Estructura esperada del archivo

- **Columna A**: Valores de débitos (formato numérico).
- **Columna B**: Valores de créditos (formato numérico).
- **Columna E**: Todos los movimientos (positivo y negativo), importados desde el extracto en PDF.
- El archivo debe tener los datos comenzando desde la fila 2 (encabezados opcionales en fila 1).

> ⚠️ Los valores en A y B deben estar en formato numérico limpio. La macro ignora texto y datos mal formateados.

---

## 🧰 Instrucciones de uso

1. Abre tu archivo Excel con los datos del extracto.
2. Presiona `Alt + F11` para abrir el Editor de VBA.
3. Importa el archivo `.bas` con esta macro.
4. Ejecuta la macro `SepararYCompararValores`.
5. Revisa las hojas generadas con los valores sin coincidencias.

---

## 🔐 Sobre los datos

Los datos utilizados en esta macro **pertenecen a una empresa privada** y no se incluyen en este repositorio por motivos de confidencialidad.  
Sin embargo, el código es completamente funcional y puede adaptarse fácilmente a cualquier otro extracto bancario con una estructura similar.

---

## 🧾 Comparación con la versión anterior

- Esta versión **no es una mejora directa**, sino una **implementación paralela** pensada para un segundo flujo de trabajo.
- Esta versión realiza también la **conciliación cruzada**, genera hojas dinámicas y convierte negativos para análisis contable.

---

## ✅ Requisitos

- Microsoft Excel 2016 o superior.
- Editor de VBA habilitado.
- Archivo de trabajo estructurado con columnas A, B y E correctamente alineadas.
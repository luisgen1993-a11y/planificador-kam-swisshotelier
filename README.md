# planificador-kam-swisshotelier
Tablero de rendimiento y comisiones

- `index.html`: el tablero (checklist diario, ventas, pipeline de setup y planilla).
- `apps-script/Code.gs`: el backend en Google Apps Script que guarda todo en Google Sheets.

## Cómo se guardan los datos

- **Primera hoja del Spreadsheet**: una fila por día con las tareas completadas.
- **Hoja `Ventas`**: una fila por venta. Se crea sola la primera vez y se llena con las
  ventas que antes estaban guardadas dentro de cada día. Si una venta aparecía en varios
  días (por ejemplo, registrada en Setup y activada después), se conserva la versión más reciente.
- **Ajustes de compensación** (salario, meta, bono, etc.): en las propiedades del script,
  editables desde el botón **Ajustes** del tablero.

## Desplegar el Apps Script

1. Abre el Spreadsheet → **Extensiones → Apps Script**.
2. Reemplaza todo el código por el contenido de `apps-script/Code.gs` y guarda.
3. **Configuración del proyecto** (ícono de engranaje) → **Propiedades de secuencia de comandos**
   → agrega la propiedad `TOKEN` con una clave larga que solo tú conozcas.
4. **Implementar → Gestionar implementaciones** → edita la implementación actual (lápiz)
   → en **Versión** elige **Nueva versión** → **Implementar**. Así la URL no cambia.
5. Abre el tablero: la primera vez te pedirá la clave del paso 3 y la recordará en ese navegador.

La clave nunca va en este repositorio, porque es público.

# 🍱 Comedor Petroceno · Control de Comidas

Aplicativo web instalable (PWA) para el control diario de comidas empacadas del
**Comedor Administrativo Petroceno**.

Funciona **sin internet**, **sin servidor** y **sin internet móvil**.

---

## 📁 Contenido

| Archivo | Para qué sirve |
|---|---|
| `index.html` | La aplicación completa (interfaz + lógica) |
| `app-xlsx.js` | Motor de Excel propio: genera el `.xlsx` con el formato original |
| `manifest.json` | Datos de la PWA (nombre, iconos, color, atajos) |
| `service-worker.js` | Permite que la app funcione sin conexión |
| `offline.html` | Pantalla que se muestra si no hay conexión |
| `icon-192.png` / `icon-512.png` / `icon-maskable.png` | Iconos para el celular |

**No hay CDN, ni librerías externas, ni `node_modules`.** Todo va incluido.

---

## 🖥️ Cómo usarlo en la computadora

**Opción 1 — Doble clic (la más fácil)**

Abre `index.html` con doble clic. Funciona todo: capturar, guardar, exportar Excel.

**Opción 2 — Servirlo (para probarlo como PWA real)**

Abre PowerShell en esta carpeta y ejecuta:

```powershell
python -m http.server 8000
```

Luego entra en `http://localhost:8000`

> Si no tienes Python, cualquier otro servidor local sirve igual.

---

## 📱 Cómo instalarlo en el Android

1. Sube la carpeta a **GitHub Pages**, **Netlify** o cualquier hosting con **HTTPS**.
2. Abre la dirección en **Chrome** del celular.
3. Menú ⋮ → **"Agregar a la pantalla de inicio"** (o "Instalar app").
4. Listo: tienes el ícono en el escritorio y **abre sin internet**.

---

## ✅ Funciones

- **Captura por área** — CCR, CCR ADM, DSI, LABORATORIO, MSOP, SALUD, PLANIFICACION, BOMBEROS (y las que agregues).
- **5 menús** — MENU 1, MENU 2, MENU 3, MENU 4, DIETA, con botón de `+` y `−` para captura rápida.
- **Total general en vivo** — se actualiza mientras capturas.
- **Personal de entrega** — nombre y cargo de quien recibe.
- **Histórico** — se guarda un snapshot automático de cada día.
- **Exportar a Excel** — con el **formato exacto** del Excel original.
- **Exportar a PDF** — para imprimir y firmar (necesita internet solo la primera vez).
- **Tema claro y oscuro** — botón 🌙 / ☀️ en la barra superior.
- **Tus datos no salen del celular** — todo se guarda en el dispositivo.

---

## 📊 El Excel que genera

Reproduce el archivo original:

- Título `SOLICITUD DE COMIDAS EMPACADAS` combined en `B1:E1`
- Subtítulo `COMEDOR ADMINISTRATIVO PETROCEDENO`
- Encabezado `FECHA | MENU 1..4 | DIETA | TOTAL`
- Una fila por área, con fórmula `=SUM(B4:F4)`
- Fila `TOTAL` con fórmulas `=SUM(...)` en cada columna
- Bloque de resumen **con fórmulas vivas** (siempre cuadra, nunca se desactualiza)
- Tipografía **Calibri**, anchos de columna y alturas de fila idénticos al original

Ventaja: si cambias una cantidad en Excel, **los totales se recalculan solos**.

---

## ⚠️ Nota sobre el PDF

El PDF usa `jsPDF`, que se descarga la **primera vez** que lo usas. El **Excel no**
necesita internet en ningún momento.

---

*Desarrollado por Rosa Elvira Velásquez · Petroceno*

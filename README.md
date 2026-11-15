# URBANSE — web oficial

Web de Urbanse, MC del sur de Buenos Aires.

- `index.html` — página para el público: videos, música, fechas, tienda y acceso a booking.
- `booking.html` — página para promotores, festivales y marcas: formatos, audiencia, trayectoria y formulario (fechas, presskit y prensa).
- `img/` — fotos optimizadas.
- `apps-script.gs` — script del Google Sheet que alimenta fechas, videos, merch, redes, pedidos de fans y solicitudes de booking. Se pega en el Sheet (Extensiones → Apps Script) y no forma parte del sitio.

## Flujo de trabajo

- Las ediciones se hacen en la rama `develop`.
- Cuando un cambio está listo, se abre un Pull Request de `develop` a `main` para revisarlo.
- `main` es la versión que se publica.

## Conectar el Sheet

1. Pegar `apps-script.gs` en el Sheet y ejecutar `setup`.
2. Implementar como aplicación web (Ejecutar como: yo · Acceso: cualquier persona) y copiar la URL `/exec`.
3. Pegar esa URL en `SHEET_API_URL` dentro de `index.html` y de `booking.html`.

Diseño y desarrollo: [Isiselinor](https://www.instagram.com/isiselinor/)

# CV Match — Diseño con Bootstrap

Esta versión utiliza Bootstrap 5.3.8 vía CDN para la estructura, cuadrícula, espaciado, botones, formularios y componentes básicos. `styles.css` contiene únicamente detalles visuales propios.

## Archivos
- `index.html`: página y enlaces CDN de Bootstrap.
- `styles.css`: detalles de marca y estilos de la ilustración.
- `app.js`: validación básica de archivo/oferta; no conecta con IA todavía.

## Probar localmente
1. Descomprime el ZIP y abre la carpeta en Visual Studio Code.
2. Abre `index.html` en el navegador o usa Live Server.
3. Se requiere conexión a internet para cargar Bootstrap desde el CDN.

## Publicar
Sube los archivos al repositorio de GitHub y conecta el repositorio con Vercel.

## Importante
Esta versión es solo la interfaz. La API de OpenAI se añadirá más adelante mediante un backend seguro en Vercel. Nunca guardes la clave de OpenAI en `app.js` ni en otro archivo público.

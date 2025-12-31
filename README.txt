Kensa Tasks — Admin (v1.8) — Split perfecto por lenguajes
===============================================================
Archivos:
- admin.html   : Solo markup. Enlaza a styles.css y app.js
- styles.css   : TODO el CSS extraído de <style> en el mismo orden del original
- app.js       : TODO el JS inline (sin 'src'), concatenado en el mismo orden

Garantías:
- No se cambiaron IDs, clases, atributos ni el contenido de CSS/JS.
- Se respetó el ORDEN de ejecución de <script> inline, por lo que las funciones mantienen
  el comportamiento idéntico al HTML original.
- El bloque "Interior del caso" y el resto del markup permanecen intactos en admin.html.

Uso:
- Abrir admin.html con el query de tenant, p.ej.: admin.html?tenant=demo-kensa
- Realizar ajustes de estilo en styles.css y de lógica en app.js.

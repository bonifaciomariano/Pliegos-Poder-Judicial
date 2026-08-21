# Pliegos del Poder Judicial en trámite

Navegador de los pliegos judiciales (jueces, fiscales, defensores, camaristas,
conjueces) que el Senado tiene en trámite ante la Comisión de Acuerdos.

**Publicado en GitHub Pages:** cada `git push` a `main` dispara un workflow
que vuelve a correr el pipeline completo (Excel/CSV/MD/PDF → JSON → HTML) y
publica el resultado. La página siempre refleja el último commit — no hay
que generar nada a mano ni acordarse de correr un build antes de subir.

➡️ Una vez que actives GitHub Pages (ver abajo), el sitio va a quedar en:
`https://bonifaciomariano.github.io/Pliegos-Poder-Judicial/`

## Activar GitHub Pages (una sola vez)

1. En el repo, andá a **Settings → Pages**.
2. En **Build and deployment → Source**, elegí **GitHub Actions**.
3. Listo. El workflow `.github/workflows/deploy.yml` ya está en el repo y se
   dispara solo en el próximo push (o corrélo manualmente desde la pestaña
   **Actions** con "Run workflow" si querés publicarlo ya, sin esperar un
   commit nuevo).

## Archivos

- `data_builder.py` — lee las fuentes y genera `pliegos_data.json`.
- `build_html.py` — embebe ese JSON (y las fuentes Montserrat) en
  `template.html` y produce `pliegos_judiciales.html`.
- `template.html` — plantilla con placeholders `__FONT400__`/`__FONT600__`/
  `__FONT700__`/`__DATA_JSON__`.
- `pliegos_judiciales.html` — navegador autocontenido (HTML/CSS/JS, sin
  frameworks). Se puede abrir directo con doble clic, no necesita servidor.
  Es lo mismo que termina publicado en GitHub Pages como `index.html`.
- `fonts/` — subset de Montserrat en `.woff2`, embebido en base64 al generar
  el HTML.
- Fuentes de datos en esta carpeta: `Acuerdos_gestion_Milei.xlsx`,
  `AYUDA_MEMORIA_2026__AC_para_dar_cuenta.csv`,
  `Audiencias_publicas_acuerdos.md`,
  `BOLETIN_DE_REUNIONES_DE_COMISIONES_91_2026.pdf`.
- `nuevas_od/` — exportaciones sueltas de Órdenes del Día nuevas (`.csv` o
  `.xlsx`, columnas `Número`/`Sobre los expedientes`/`Periodo`/`Fecha
  Dictamen`) que todavía no están volcadas en el Excel madre. Se pueden ir
  acumulando ahí, cada corrida las relee todas.

## Cómo actualizar los datos

**El flujo normal es simplemente subir el archivo nuevo y pushear** — el
workflow de Pages hace el resto. Localmente, para probar antes de subir:

1. Reemplazá el archivo que corresponda en esta carpeta (mismo nombre) o
   soltá un `.csv`/`.xlsx` nuevo en `nuevas_od/`.
2. Corré:

   ```bash
   pip install -r requirements.txt   # openpyxl y pdfplumber, si hace falta
   python3 data_builder.py
   ```

   Esto regenera `pliegos_data.json` y además imprime en consola:
   - los expedientes excluidos por ser administrativos o retiros de mensaje,
   - los que no matchean el patrón de designación esperado (para revisar a mano),
   - los que tienen un cargo no judicial (si aparece alguno),
   - avisos si el ayuda-memoria de "dar cuenta" no coincide con la planilla,
   - avisos de Orden del Día con formato raro.

3. Volvé a generar el HTML con el JSON actualizado:

   ```bash
   python3 build_html.py
   ```

   Abrilo en el navegador para revisar antes de commitear.

4. `git add -A && git commit -m "..." && git push` — el push ya deja todo
   publicado en un par de minutos (podés seguirlo en la pestaña **Actions**
   del repo).

Si el Excel cambia de nombre o ubicación, pasale la ruta explícita:

```bash
python3 data_builder.py --xlsx "/ruta/a/Acuerdos_nuevo.xlsx" --out pliegos_data.json
```

También podés apuntar `--csv`, `--md`, `--pdf` y `--nuevas-od-dir` a rutas
distintas de la misma forma.

## Notas sobre la lógica de datos

- Solo se procesan filas con `TIPO == 'AC'`.
- El hipervínculo real al expediente se lee de la fórmula
  `=HYPERLINK(...)` de la columna B (por eso `openpyxl` se abre con
  `data_only=False`).
- La fecha de audiencia pública usa como fuente primaria el archivo
  `Audiencias_publicas_acuerdos.md` (más confiable); si un expediente no
  aparece ahí, se usa `FECHA_EGRESO1` / `FECHA INGRESO DICTAMEN` de la
  planilla (son el mismo valor, se verificó por script) — siempre que sea
  una fecha real y no un placeholder tipo `" -"`.
- Categorías, en orden de prioridad: `dar_cuenta` → `sancionado` →
  `con_od` → `sin_audiencia`.
- Los mensajes puramente administrativos (p. ej. "asigna salas y
  vocalías") y los retiros de mensaje se excluyen del navegador y se
  listan en el log de la consola.
- `AUDIENCIAS_PROGRAMADAS_MANUAL` (en `data_builder.py`) permite cargar a
  mano un aviso de audiencia pública programada cuando llega como texto
  suelto en vez de un boletín PDF nuevo.
- `AUDIENCIAS_PROGRAMADAS_VENCIDAS` (en `data_builder.py`) apaga el badge
  "programada" de fechas que ya pasaron; hay que sumarlas ahí a mano y
  agregar la fecha real como "realizada" en `Audiencias_publicas_acuerdos.md`.
- El badge "Audiencia realizada, dictamen pendiente" se calcula solo (no
  hace falta tocar nada): aparece en cualquier pliego que ya tuvo audiencia
  pero todavía sigue en categoría `sin_audiencia`, y se apaga automáticamente
  apenas consigue Orden del Día.

# litigiar-valores

Scraper automatizado de valores oficiales para la extensión **LitigiAR**.
Corre en GitHub Actions cada 6 horas y actualiza `valores.json`.

## Fuentes

| Valor | Fuente | Frecuencia |
|---|---|---|
| IPC (mensual, interanual, acumulada) | INDEC vía datos.gob.ar | Mensual |
| UVA / CER / UVI | BCRA vía datos.gob.ar | Diaria |
| ICL | BCRA `diar_icl.xls` | Diaria |
| RIPTE | MTSS vía datos.gob.ar | Mensual |
| SMVM | CNEPySMVyM vía datos.gob.ar | Bimestral |
| JUS Ley 14.967 / DL 8904 / Previsional | ColProBA + Colegio La Matanza + Morón | ~Trimestral |
| Bono Ley 8480 (verde/azul) | ColProBA + colegios departamentales | ~Semestral |
| Canasta Crianza | INDEC XLSX | Mensual |

## Setup (una sola vez)

```bash
cd litigiar-valores
git init
git branch -M main
git add .
git commit -m "init: scraper de valores oficiales"
# Crear repo en GitHub (privado o público)
git remote add origin git@github.com:TU_USUARIO/litigiar-valores.git
git push -u origin main
```

Luego activar GitHub Actions: **Settings → Actions → General → Allow all actions**.

En **Settings → Actions → General → Workflow permissions** seleccionar **"Read and write permissions"** (necesario para que el bot pueda commitear `valores.json`).

## Probar localmente

```bash
npm install
npm run scrape
# Inspeccionar valores.json
```

## URL pública del JSON (una vez publicado)

```
https://raw.githubusercontent.com/TU_USUARIO/litigiar-valores/main/valores.json
```

Este URL se configura en `al-valores.js` del extension como `BACKEND_URL`.

## Monitoreo

- Cada ejecución queda registrada en **Actions** tab
- Si el scraper rompe, GitHub envía email al mantenedor
- Cada actualización produce un commit con fecha/hora (audit trail)

## Política de fallos (tolerante)

| Situación | Comportamiento |
|---|---|
| Scraper OK | `status: "ok"` → chip verde |
| Scraper falla pero hay valor previo válido | `status: "stale"` → se mantiene el último conocido |
| Scraper falla y nunca hubo valor | `status: "fallback"` → valor hardcodeado |
| >5 fallbacks simultáneos | Workflow falla → alerta por email |

## Agregar fuentes redundantes

Si una fuente (ej. ColProBA) cambia su HTML y rompe:
1. El scraper devuelve `stale` para JUS/Bonos (mantiene último valor OK)
2. Se puede agregar otro colegio departamental (Lomas, Avellaneda-Lanús, etc.) como fallback en `scripts/scrape.js` función `refreshPBA`
3. Los regex en `scrapearHTMLcolegio` son flexibles ante typos comunes

## Licencia

Privado — uso interno de LitigiAR.

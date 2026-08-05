# Integració de la web pública (PVP) amb la calculadora

Aquest document descriu com una web pública de cara al client final (p. ex. una
app Astro SSR) consumeix els preus **PVP** de la calculadora, sense veure mai
costos ni marges. La calculadora és la **font de veritat**: els preus es
calculen en viu a cada petició, així que no cal exportar ni sincronitzar res.

## Model de seguretat

- Autenticació per **token secret dedicat** a la capçalera HTTP `X-Pvp-Token`.
- El token és **independent** del bridge de reusrevela-web, per poder-lo revocar
  a part. Suporta rotació sense downtime.
- ⚠️ El token **només pot viure al servidor** de la web pública (una API route /
  funció de servidor). **Mai** ha d'arribar al navegador. A Astro, això vol dir
  que **no** pot tenir el prefix `PUBLIC_`.

### Variables d'entorn (a la calculadora)

| Variable | Ús |
|---|---|
| `PUBLIC_PVP_TOKEN` | Token vàlid actual. |
| `PUBLIC_PVP_TOKEN_NEXT` | (Opcional) token nou durant una rotació; quan la web ja envia el nou, es promou a `PUBLIC_PVP_TOKEN` i s'esborra aquest. |
| `BRIDGE_ALLOWED_IPS` | (Opcional, ja existent) allowlist d'IPs que també aplica a aquests endpoints. |

### Variable d'entorn (a la web pública / Astro, a Railway)

```
PRICING_URL=https://calculadora.reusrevela.cat
PRICING_TOKEN=<el mateix valor que PUBLIC_PVP_TOKEN>   # SENSE prefix PUBLIC_
```

## Endpoint disponible (Fase 1)

### `GET /api/public/pvp/compute`

Retorna el **preu PVP** (client final) d'un producte concret: aplica el marge
retail per defecte (`marge_defecte`, avui 60%) sobre el PVD i hi suma l'IVA.
**Mai** retorna cost ni PVD.

**Capçalera:** `X-Pvp-Token: <token>`

**Query params** (idèntics a `/api/public/compute`):

| Param | Obligatori | Valors |
|---|---|---|
| `kind` | sí | `impressio` · `laminate` · `protter` · `frame` |
| `width_cm` | sí | enter/decimal > 0 |
| `height_cm` | sí | enter/decimal > 0 |
| `qty` | no | enter ≥ 1 (defecte 1) |
| `paper` | no | impressió: `lustre`, `silk`, `fine_art`… |
| `finish` | no | `none` (defecte) · `laminate` · `protter` · `foam` |
| `moldura_id` | només si `kind=frame` | referència de la motllura |

**Resposta 200:**

```json
{
  "ok": true,
  "kind": "frame",
  "width_cm": 30, "height_cm": 40, "qty": 2,
  "marge_pct": 60.0,
  "vat_rate": 0.21,
  "pvp_net": 32.00,
  "iva": 6.72,
  "pvp_total": 38.72
}
```

- `pvp_net` = preu al client **sense IVA**.
- `pvp_total` = preu al client **amb IVA inclòs** (el que sol veure el client).

**Errors:** `403 forbidden` (token dolent/absent) · `400 unknown_kind` ·
`400 invalid_size` · `400 missing_moldura_id` · `404 impressio_not_found` ·
`404 laminate_not_found` · `404 moldura_not_found` · `500 compute_failed`.

## Exemple d'integració (Astro SSR)

La regla d'or: **el navegador parla només amb la teva pròpia API route**, que és
qui afegeix el token i crida la calculadora.

`src/pages/api/pricing.ts`:

```ts
import type { APIRoute } from "astro";

export const prerender = false;

const PRICING_URL = import.meta.env.PRICING_URL;
const PRICING_TOKEN = import.meta.env.PRICING_TOKEN; // sense prefix PUBLIC_

export const GET: APIRoute = async ({ url }) => {
  const qs = url.searchParams; // kind, width_cm, height_cm, qty, moldura_id, finish...
  const upstream = new URL("/api/public/pvp/compute", PRICING_URL);
  upstream.search = qs.toString();

  const res = await fetch(upstream, {
    headers: { "X-Pvp-Token": PRICING_TOKEN },
  });
  const data = await res.json();

  // Reenviem tal qual (només conté PVP net/iva/total, mai cost).
  return new Response(JSON.stringify(data), {
    status: res.status,
    headers: { "content-type": "application/json" },
  });
};
```

Al navegador:

```js
const r = await fetch(`/api/pricing?kind=frame&width_cm=30&height_cm=40&qty=2&moldura_id=${encodeURIComponent(ref)}`);
const p = await r.json();
if (p.ok) mostrarPreu(p.pvp_total); // preu amb IVA al client
```

## Notes de preu

- El PVP surt de `PVD × (1 + marge_defecte/100)` + IVA 21%, la mateixa convenció
  que ja fa servir reusrevela-web. La calculadora, a la seva pantalla, pot
  mostrar per a la **impressió** un PVP lleugerament diferent perquè hi aplica
  marges per tram d'àrea; si vols que la web pública els respecti exactament,
  es pot afegir un marge retail per categoria (impressió, marcs…) sense canviar
  el contracte de l'API.

## Fase 2 (pendent)

- `POST /api/public/order` — recollir la comanda/pressupost del client i desar-la
  a `comandes` (estat "nou"), retornant un número de referència.
- Catàleg públic (llista de productes/opcions per muntar el selector).
- Suport de **llenç** (canvas) al motor de preu públic.

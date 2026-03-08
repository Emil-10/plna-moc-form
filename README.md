# Plná moc - form

Jednoduchá webová aplikace pro vyplnění plné moci na převod vozidla.

## Funkce

- našeptávání firmy/osoby přes ARES
- automatické doplnění podle IČO
- výběr rozsahu zplnomocnění
- živý náhled dokumentu
- export do PDF

## Lokální spuštění

Pro základní formulář stačí otevřít `index.html` v prohlížeči, nebo spustit jednoduchý server:

```bash
python -m http.server 4399
```

Pak otevřít:

```text
http://127.0.0.1:4399/
```

Poznámka: režim přes jednoduchý statický server nepodporuje Cloudflare Pages Functions, takže nebude fungovat VIN lookup přes `/api/vin`.

Pro plnou funkčnost včetně VIN registru spusťte aplikaci přes Pages runtime:

```bash
npx wrangler pages dev . --port 8788
```

Pak otevřít:

```text
http://127.0.0.1:8788/
```

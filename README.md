# Link Cards WebPart

- **Řešení:** link-cards-webpart-client-side-solution
- **Webpart:** LinkCardsWebPart
- **Solution GUID:** `ae20369c-020a-4804-b0c6-8165e55b7455`
- **Component GUID:** `ca899b6b-69c0-4c40-8986-c7b5fef99156`
- **SPFx:** 1.22.1
- **Node.js:** 22.14+ a nižší než 23

## Build

```bash
npm install
npm run build
npm run package-solution
```

Výstup: `sharepoint/solution/link-cards-webpart.sppkg`

Karty se spravují přímo ve webpartu v režimu úprav stránky. Lze je přidávat, odebírat, přesouvat, vybírat obrázky z knihovny Site Assets (Prostředky webu), vybírat cílové stránky z Site Pages (Stránky webu) a nastavit barvu celé karty.

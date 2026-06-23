# BC SaaS Roadmap

Naujinimų planavimo ir sekimo įrankis Business Central SaaS klientams.

## Funkcionalumas

- Naujinimų roadmap su Gantt grafiku
- BC aplinkų versijų sekimas (automatiškai iš BC Admin API)
- App Registration secrets galiojimo stebėjimas (automatiškai iš Microsoft Graph)
- GDAP / Admin Relationships stebėjimas (automatiškai)
- Testavimo laikotarpių planavimas (rankinis įvedimas → SharePoint)
- Klientų informavimo terminų sekimas
- Statusų valdymas

## Architektūra

```
Azure Static Web Apps (frontend)
        ↓
Azure Functions (backend)
    ↙           ↘
Microsoft API    SharePoint Lists
```

## Diegimas

Žiūrėkite `docs/DIEGIMO_INSTRUKCIJA.md`

## Struktūra

```
bc-roadmap/
├── frontend/          # HTML/CSS/JS web aplikacija
│   ├── index.html
│   ├── style.css
│   ├── app.js
│   └── staticwebapp.config.json
├── backend/           # Azure Functions
│   ├── functions/
│   │   ├── shared/graph.js     # Microsoft API helper
│   │   ├── getClients/
│   │   ├── getUpdates/
│   │   ├── saveUpdate/
│   │   ├── deleteUpdate/
│   │   ├── getSecrets/
│   │   ├── saveSecret/
│   │   ├── deleteSecret/
│   │   └── getVersions/
│   ├── host.json
│   ├── package.json
│   └── local.settings.json     # NE commitinti į git!
├── docs/
│   └── DIEGIMO_INSTRUKCIJA.md
└── .github/workflows/deploy.yml
```

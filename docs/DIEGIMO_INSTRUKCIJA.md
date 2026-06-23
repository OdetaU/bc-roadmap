# BC SaaS Roadmap — Diegimo instrukcija

## Apžvalga

Ši aplikacija susideda iš:
- **Frontend** — HTML/CSS/JS (veikia naršyklėje)
- **Backend** — Azure Functions (Node.js, jungiasi prie Microsoft API)
- **Duomenų bazė** — SharePoint Lists (naujinimų planai, pastabos)
- **Automatiniai duomenys** — BC Admin API, Microsoft Graph (versijos, secrets, GDAP)

---

## 1 žingsnis — App Registration Azure portale

1. Eikite į https://portal.azure.com
2. Ieškokite **"App registrations"** → **New registration**
3. Užpildykite:
   - Name: `BC-Roadmap-App`
   - Supported account types: `Accounts in this organizational directory only`
   - Redirect URI: palikite tuščią (užpildysime vėliau)
4. Spaudžiate **Register**
5. **Išsaugokite** šiuos duomenis (jų reikės vėliau):
   - `Application (client) ID` → tai yra jūsų `CLIENT_ID`
   - `Directory (tenant) ID` → tai yra jūsų `TENANT_ID`

### 1.1 — Sukurkite Client Secret

1. Kairėje meniu: **Certificates & secrets** → **New client secret**
2. Description: `BC Roadmap Secret`
3. Expires: **24 months**
4. Spaudžiate **Add**
5. **Iš karto nukopijuokite Value** — vėliau jo nebematysite!
   - Tai yra jūsų `CLIENT_SECRET`

### 1.2 — Suteikite API teises

1. Kairėje: **API permissions** → **Add a permission**

**Microsoft Graph:**
- Application permissions → `Application.Read.All`
- Application permissions → `DelegatedAdminRelationship.Read.All`

**Business Central:**
- Spaudžiate **APIs my organization uses**
- Ieškote `Dynamics 365 Business Central`
- Application permissions → `API.ReadWrite.All`

2. Spaudžiate **Grant admin consent for StrongPoint**

---

## 2 žingsnis — SharePoint sąrašų sukūrimas

Eikite į savo SharePoint svetainę ir sukurkite 3 sąrašus:

### Sąrašas: BC_Updates
| Stulpelio pavadinimas | Tipas |
|---|---|
| ClientId | Single line of text |
| CurrentVersion | Single line of text |
| NewVersion | Single line of text |
| Environment | Single line of text |
| DateTest | Date and Time |
| DateTestEnd | Date and Time |
| DateProd | Date and Time |
| DateNotify | Date and Time |
| Status | Single line of text |
| Notes | Multiple lines of text |

### Sąrašas: BC_Secrets
| Stulpelio pavadinimas | Tipas |
|---|---|
| ClientId | Single line of text |
| Type | Single line of text |
| Name | Single line of text |
| ExpiresAt | Date and Time |
| WarnDays | Number |
| Notes | Multiple lines of text |
| Source | Single line of text |
| SecretId | Single line of text |
| RelationshipId | Single line of text |

### Sąrašas: BC_Clients
| Stulpelio pavadinimas | Tipas |
|---|---|
| TenantId | Single line of text |
| Contact | Single line of text |
| Email | Single line of text |
| BcVersion | Single line of text |
| Notes | Multiple lines of text |

### Kaip sužinoti SharePoint Site ID

Atidarykite naršyklėje:
```
https://uabstrongpoint.sharepoint.com/sites/JUSU_SVETAINE/_api/site/id
```
Gausite XML su `<d:Id>` — tai ir yra jūsų `SP_SITE_ID`.

---

## 3 žingsnis — GitHub repozitorija

1. Sukurkite naują repozitoriją GitHub (pvz. `bc-roadmap`)
2. Įkelkite visą projekto kodą:

```bash
cd bc-roadmap
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/JUSU_ORGANIZACIJA/bc-roadmap.git
git push -u origin main
```

---

## 4 žingsnis — Azure Static Web Apps sukūrimas

1. https://portal.azure.com → **Create a resource** → **Static Web App**
2. Užpildykite:
   - Resource group: sukurkite naują arba naudokite esamą
   - Name: `bc-roadmap`
   - Plan type: **Free**
   - Region: `West Europe`
   - Source: **GitHub**
3. Prisijunkite prie GitHub ir pasirinkite savo repozitoriją
4. Build details:
   - App location: `/frontend`
   - Api location: `/backend`
   - Output location: (palikite tuščią)
5. Spaudžiate **Review + create** → **Create**

Azure automatiškai sukurs GitHub Actions workflow ir pradės pirmą deployment'ą.

---

## 5 žingsnis — Environment Variables konfigūracija

Azure portale, jūsų Static Web App:
1. **Configuration** → **Application settings**
2. Pridėkite šiuos kintamuosius:

| Name | Value |
|---|---|
| TENANT_ID | (iš 1 žingsnio) |
| CLIENT_ID | (iš 1 žingsnio) |
| CLIENT_SECRET | (iš 1 žingsnio) |
| SP_SITE_ID | (iš 2 žingsnio) |
| SP_SITE_URL | https://uabstrongpoint.sharepoint.com/sites/JUSU_SVETAINE |

3. **Save**

---

## 6 žingsnis — Frontend App Registration konfigūracija

Frontend taip pat reikia atskiro App Registration (arba galima naudoti tą patį):

1. App Registration → **Authentication** → **Add a platform** → **Single-page application**
2. Redirect URI: `https://JUSU-APP.azurestaticapps.net`
3. Spaudžiate **Configure**

Tada faile `frontend/app.js` pakeiskite:
```javascript
clientId: 'ĮRAŠYKITE_FRONTEND_APP_CLIENT_ID',
```

---

## 7 žingsnis — GitHub Secret

GitHub Actions reikia žinoti Azure deployment token:

1. Azure portale, jūsų Static Web App → **Manage deployment token**
2. Nukopijuokite token
3. GitHub repozitorijoje → **Settings** → **Secrets and variables** → **Actions**
4. **New repository secret**:
   - Name: `AZURE_STATIC_WEB_APPS_API_TOKEN`
   - Value: (nukopijuotas token)

---

## Baigta!

Po pirmojo push į `main` branch, GitHub Actions automatiškai:
1. Sudeploys frontend į Azure Static Web Apps
2. Sudeploys backend (Azure Functions) tą pačią aplikaciją

Jūsų aplikacija bus prieinama adresu, kurį suteiks Azure (pvz. `https://bc-roadmap-xxxx.azurestaticapps.net`).

---

## Pastabos dėl duomenų

| Duomenų šaltinis | Automatinis | Rankinis |
|---|---|---|
| Klientų sąrašas (iš BC Admin) | ✓ | |
| BC aplinkų versijos | ✓ | |
| App Registration secrets galiojimas | ✓ | |
| GDAP relationships | ✓ | |
| Testavimo datos ir planai | | ✓ (SharePoint) |
| Naujinimų statusas | | ✓ (SharePoint) |
| Pastabos | | ✓ (SharePoint) |
| Klientų kontaktai | | ✓ (SharePoint) |

---

## Pagalba

Jei iškyla klausimų diegimo metu — kiekvienas žingsnis gali būti atliktas per 5-15 minučių.
Pilnas diegimas pirmą kartą: ~2-3 valandos.

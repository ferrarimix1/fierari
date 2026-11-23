# Facturi Fierari: fierari.avrusu.eu

The house's utility bills, split between the four apartments. It is a single web page that phones can install like an app.

- **The front page opens straight away, with no password.** Anyone with the address sees the latest month to pay:
  - each apartment's amount, with its part of every bill;
  - the total of each bill;
  - how each bill was split, using that month's numbers (floor-area shares, people, meter readings, price per kWh);
  - at the bottom, the outside meters (the providers' gas meter and the two electricity meters): the latest readings and the use since the ones before. They are not part of the split.

  The arrows next to the month go to earlier months. Nothing can be changed there.
- **Changes are made under Administrare** (the link at the bottom, or https://fierari.avrusu.eu/#admin). That is the full app: adding and editing months, the history chart, the Excel download and the settings.
- **Meter readings with the password:** "Trimite citirile" at the bottom (https://fierari.avrusu.eu/#citiri). Whoever has the password can type the five meter readings, and optionally the three outside meters (gas, Electric Sus, Electric Jos). They show on the front page as new readings, with the use so far. The owner's form for the next month starts from them.
- **Only the owner can change anything.** A device can add or edit months once it is connected with a GitHub key that may write to this repository. The key stays on that device. Each save is one commit to `data.json`, and the site shows it to everyone within 1–2 minutes.
- **The figures are public.** `data.json` holds the months in plain text, in this public repository and on the site. Search engines are asked not to index the site, but anyone with the address can read it. Keep names and other personal details out of the notes.

## One-time setup

1. **Repository.** On GitHub, create a new **public** repository named exactly `fierari`. Upload every file from this folder with "Add file → Upload files".
2. **GitHub Pages.** Go to Settings → Pages → Build and deployment. Set Source to "Deploy from a branch", branch `main`, folder `/ (root)`, then Save. Under Custom domain, enter `fierari.avrusu.eu` and Save.
3. **DNS at Netim.** In the DNS zone of `avrusu.eu`, add a record:
   - type `CNAME`
   - name/host `fierari`
   - value `ferrarimix1.github.io`

   It can take from a few minutes to a few hours to work.
4. **HTTPS.** Back in Settings → Pages, once the DNS check is green, tick **Enforce HTTPS**.
5. **The key for adding months.** Go to GitHub → Settings → Developer settings → Personal access tokens → Fine-grained tokens → Generate new token:
   - Repository access: *Only select repositories* → `fierari`
   - Permissions → Repository permissions → **Contents: Read and write**
   - Expiration: up to one year

   Copy the key; GitHub shows it only once.
6. **First time on the site.** Open https://fierari.avrusu.eu and press **Administrare** at the bottom. Paste the key, choose a password (twice), then press **Conectează și salvează parola**.
   - The key is checked, then kept on the site in `panel.json`, locked with that password (AES-GCM, key from PBKDF2-SHA256).
   - From then on, on any device, the password is all that is needed: for Administrare and for "Trimite citirile".
   - Anyone who has the password can change anything that key reaches, so give it only to people you trust.
   - Under Administrare → Cheia GitHub you can change the password or switch it off (`panel.json` becomes `{}`).
   - When the key expires, use "Am o cheie GitHub nouă" in Administrare.

Optional: under GitHub → Settings → Pages, *verify* `avrusu.eu`, so nobody else can ever attach your domain to their own GitHub site.

## Day to day

- **Adding a month:** go to Administrare and press "Adaugă …". DIGI (98) and rubbish (22 per apartment) are filled in, and so are the meter readings (and the outside meters) if they were sent through "Trimite citirile". Type gas, water and the people.
  - Electricity: one field, the price per kWh. It starts at the latest price; to begin with that is November 2026's, 1,283.52 lei ÷ 1,156 kWh = 1.110311 lei/kWh. Change it when the price changes.
  - You can type the price (1,11) or the division from the bill (1099,43/990). It is kept to 6 decimals, and the line under it shows what the month's kWh come to.
  - Afterwards, "Facturile de plată" at the top goes back to the front page.
- **Tenants:** send them the address. They can add the page to the home screen (Share → Add to Home Screen).
- **Key expired or a phone lost:** delete the key on GitHub (Settings → Developer settings → tokens), make a new one, and use "Am o cheie GitHub nouă" under Administrare. "Cheia GitHub" → "Deconectează acest dispozitiv" makes a device forget the key and the password.

## How the bills are split

- **Gas:** by floor-area share (22.38 / 19.83 / 28.23 / 29.56 %).
- **DIGI:** split equally between the four apartments (98 lei is 24.50 each).
- **Water:** by the number of people that month.
- **Rubbish:** the same fixed amount per apartment.
- **Electricity:** each apartment pays the use on its own meter, and only the common meter's use is split, by floor-area share. Every kWh is paid at that month's price per kWh, from the electricity bill.

Amounts are rounded to the ban so the four always add up exactly to the month's total. On the front page, each apartment's part of each bill is rounded too:
- the parts of a bill add up to the bill;
- an apartment's parts add up to what it pays;
- each part stays within one ban of its exact amount, as close as those totals allow.

## Files

| File | What it is |
|---|---|
| `index.html`, `app.js`, `style.css` | The page |
| `data.json` | The months. Changed by the page itself when you save |
| `citiri.json` | The latest readings sent through the panel (`{}` when none) |
| `panel.json` | The key, locked with the password (`{}` when the password is off) |
| `fonts/`, `fonts.css` | Archivo, Inter, IBM Plex Mono (SIL Open Font License), served from here: no Google requests |
| `vendor/xlsx.full.min.js` | SheetJS (Apache-2.0), loaded only for the Excel download |
| `manifest.webmanifest`, `icon-*.png` | Lets phones install the page like an app |
| `CNAME`, `.nojekyll`, `robots.txt` | GitHub Pages settings; search engines are asked not to index the site |

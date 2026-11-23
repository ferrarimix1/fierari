# Facturi Fierari: fierari.avrusu.eu

The house's utility bills, split between the four apartments. It is a single web page that phones can install like an app.

- **The front page only shows.** After the password, everyone sees the latest month to pay:
  - each apartment's amount, with its part of every bill;
  - the total of each bill;
  - how each bill was split, using that month's numbers (floor-area shares, people, meter readings, price per kWh).

  The arrows next to the month go to earlier months. Nothing can be changed there.
- **Changes are made under Administrare** (the link at the bottom, or https://fierari.avrusu.eu/#admin). That is the full app: adding and editing months, the history chart, the Excel download and the settings.
- **Viewing needs a password.** The figures are stored only in `data.enc.json`, encrypted (AES-256-GCM, key derived with PBKDF2-SHA256, 310,000 iterations). This repository is public, and without the password the file is unreadable.
- **Only the owner can change anything.** A device can add or edit months once it is connected with a GitHub key that may write to this repository. The key stays on that device. Each save is one commit to `data.enc.json`, and the site shows it to everyone within 1–2 minutes.

## One-time setup

1. **Repository.** On GitHub, create a new **public** repository named exactly `fierari`. Upload every file from this folder with "Add file → Upload files".
2. **GitHub Pages.** Go to Settings → Pages → Build and deployment. Set Source to "Deploy from a branch", branch `main`, folder `/ (root)`, then Save. Under Custom domain, enter `fierari.avrusu.eu` and Save.
3. **DNS at Netim.** In the DNS zone of `avrusu.eu`, add a record:
   - type `CNAME`
   - name/host `fierari`
   - value `ferrarimix1.github.io`

   It can take from a few minutes to a few hours to work.
4. **HTTPS.** Back in Settings → Pages, once the DNS check is green, tick **Enforce HTTPS**. The page needs https: without it the browser refuses to decrypt.
5. **The key for adding months.** Go to GitHub → Settings → Developer settings → Personal access tokens → Fine-grained tokens → Generate new token:
   - Repository access: *Only select repositories* → `fierari`
   - Permissions → Repository permissions → **Contents: Read and write**
   - Expiration: up to one year

   Copy the key; GitHub shows it only once.
6. **Connect your device.** Open https://fierari.avrusu.eu and enter the password. Press **Administrare** at the bottom; the page asks for the key. Paste it, then press **Conectează**. Repeat this on each device you want to add months from.

Optional: under GitHub → Settings → Pages, *verify* `avrusu.eu`, so nobody else can ever attach your domain to their own GitHub site.

## Day to day

- **Adding a month:** go to Administrare and press "Adaugă …". DIGI (98) and rubbish (22 per apartment) are filled in; type gas, water, each electricity bill (total and kWh), the meter readings and the people. Afterwards, "Facturile de plată" at the top goes back to the front page.
- **Tenants:** send them the address and the password. They can tick "Ține minte pe acest dispozitiv" and add the page to the home screen (Share → Add to Home Screen).
- **Changing the password:** go to Administrare → Cheie și parolă → Parola de acces. The data file is re-encrypted; give the tenants the new password.
- **Keep the password safe.** Without it nobody can read the data, you included. The "Descarcă Excel" button gives you a readable copy at any time.
- **Key expired or a phone lost:** delete the key on GitHub (Settings → Developer settings → tokens), make a new one and connect again. "Deconectează acest dispozitiv" (in Administrare) removes the key from a device.

## How the bills are split

- **Gas:** by floor-area share (22.38 / 19.83 / 28.23 / 29.56 %).
- **DIGI and water:** by the number of people that month.
- **Rubbish:** the same fixed amount per apartment.
- **Electricity:** each apartment pays the use on its own meter. Only the common meter's use is split, by floor-area share. Every kWh is paid at that month's price, which is the total lei on the electricity bills ÷ the total kWh they bill.

Amounts are rounded to the ban so the four always add up exactly to the month's total. On the front page, each apartment's part of each bill is rounded too: the parts of a bill add up to the bill, and an apartment's parts add up to what it pays. Each part stays within one ban of its exact amount, as close as those totals allow.

## Files

| File | What it is |
|---|---|
| `index.html`, `app.js`, `style.css` | The page |
| `data.enc.json` | The data, encrypted. Changed by the page itself when you save |
| `fonts/`, `fonts.css` | Archivo, Inter, IBM Plex Mono (SIL Open Font License), served from here: no Google requests |
| `vendor/xlsx.full.min.js` | SheetJS (Apache-2.0), loaded only for the Excel download |
| `manifest.webmanifest`, `icon-*.png` | Lets phones install the page like an app |
| `CNAME`, `.nojekyll`, `robots.txt` | GitHub Pages settings; search engines are asked not to index the site |

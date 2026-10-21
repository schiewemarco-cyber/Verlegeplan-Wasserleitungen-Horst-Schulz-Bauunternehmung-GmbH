# Verlegeplan-Editor: Cloud-Anbindung (Microsoft 365) – Einrichtung durch die IT

**Ziel:** Mitarbeitende melden sich in der App mit ihrem Firmenkonto an und speichern ihre Verlegepläne in OneDrive oder in einer SharePoint-Bibliothek. Dort sind sie auf allen Geräten (PC, Tablet, Handy) verfügbar.

Die App ist eine reine Browser-Anwendung (statische Dateien, kein eigener Server). Sie spricht direkt mit Microsoft Graph. Es wird nichts bei Dritten gespeichert.

## 1. Was die IT einrichtet (ca. 15–30 Minuten)

### 1.1 App-Registrierung in Microsoft Entra ID
1. Entra Admin Center (entra.microsoft.com) → *Identität* → *Anwendungen* → *App-Registrierungen* → *Neue Registrierung*.
2. **Name:** z. B. „Verlegeplan-Editor“.
3. **Unterstützte Kontotypen:** „Nur Konten in diesem Organisationsverzeichnis“ (einzelner Mandant).
4. **Umleitungs-URI:** Plattform **„Single-Page-Anwendung (SPA)“** wählen und die Adresse der App eintragen, **genau mit abschließendem Schrägstrich**, z. B.
   `https://NAME.github.io/verlegeplan/`
   (bei Betrieb auf dem Firmenserver: dessen Adresse, z. B. `https://verlegeplan.firma.de/`).
   Es darf nur die Plattform „SPA“ verwendet werden (nicht „Web“), sonst schlägt die Anmeldung fehl.
5. Nach dem Anlegen notieren: **Anwendungs-(Client-)ID** und **Verzeichnis-(Mandanten-)ID**.
   Diese beiden Werte sind keine Geheimnisse (öffentlicher Client), es gibt **kein Client Secret**.

### 1.2 API-Berechtigungen (Microsoft Graph, *Delegiert*)
| Variante | Berechtigung | Wirkung |
|---|---|---|
| **A1: OneDrive** (jede Person speichert in ihrem eigenen OneDrive) | `Files.ReadWrite` | Zugriff nur auf die eigenen Dateien der angemeldeten Person |
| **A2: SharePoint** (gemeinsame Bibliothek, alle Berechtigten sehen alle Pläne) | `Files.ReadWrite.All` | Zugriff auf alle Dateien, auf die die angemeldete Person Zugriff hat |

- Dazu kommen die Standardberechtigungen `User.Read`, `openid`, `profile`, `offline_access` (werden von der Anmeldung automatisch angefordert).
- Je nach Einstellung des Mandanten (Benutzereinwilligung) kann eine **Administratorzustimmung** nötig sein: *API-Berechtigungen* → „Administratorzustimmung für <Mandant> erteilen“. Im Zweifel erteilen.
- *Hinweis A2:* `Files.ReadWrite.All` ist breit gefasst. Wenn das nicht gewünscht ist, kann stattdessen mit `Sites.Selected` gearbeitet werden (Zugriff nur auf eine freigegebene Site). Das erfordert eine zusätzliche Freigabe der App auf der Site per Graph-Aufruf und die Eintragung der Berechtigung unter `scopes` in `config.js`. Sprecht das vorab mit uns ab.

### 1.3 Nur Variante A2: SharePoint vorbereiten
1. Site und Dokumentenbibliothek festlegen, z. B. Site `/sites/Tiefbau`, Bibliothek „Dokumente“.
2. Die berechtigte Gruppe (Mitglieder) mit **Bearbeiten**-Recht auf die Bibliothek ausstatten.
3. Der Ordner für die Pläne (Standard: `Verlegeplaene`) wird von der App bei Bedarf selbst angelegt.

### 1.4 Zugriffsrichtlinien prüfen
- Falls *Bedingter Zugriff* gilt (verwaltete Geräte, MFA, Standorte), muss die App-Registrierung davon erfasst und die Nutzung auf Handy und Tablet im Browser erlaubt sein.
- Die Anmeldung erfolgt per Umleitung zu login.microsoftonline.com und zurück zur App-Adresse (keine Popups).

## 2. Was in der App eingetragen wird
In der Datei **`config.js`** (liegt neben `index.html`) stehen die Werte. Bitte ausfüllen und die Datei in das Repository bzw. auf den Server hochladen:

```js
window.VP_CLOUD = {
  clientId: 'HIER-DIE-CLIENT-ID',
  tenantId: 'HIER-DIE-MANDANTEN-ID',
  mode: 'onedrive',            // oder 'sharepoint'
  folder: 'Verlegeplaene',
  site: { host: '', path: '', library: '' },   // nur bei 'sharepoint'
  redirectUri: '',             // leer lassen, wenn identisch mit der App-Adresse
  scopes: []                   // leer = Standard je nach Variante
};
```
Beispiel SharePoint: `mode: 'sharepoint', site: { host: 'firma.sharepoint.com', path: '/sites/Tiefbau', library: 'Dokumente' }`

Solange `clientId` leer ist, bleibt die Cloud-Funktion unsichtbar. Die App arbeitet dann nur lokal.

## 3. Anmelde-Bibliothek (MSAL)
Die App lädt beim Anmelden die Microsoft-Bibliothek `@azure/msal-browser` in der festen Version **2.37.1**:
- Standard: von `cdn.jsdelivr.net`.
- Empfohlen für Firmenbetrieb: die Datei `lib/msal-browser.min.js` aus dem npm-Paket `@azure/msal-browser@2.37.1` herunterladen und als **`msal-browser.min.js`** neben `index.html` ablegen. Die App nutzt dann diese lokale Kopie und braucht keine externe Quelle.
Bitte prüft vor Produktivbetrieb, ob für diese Version bekannte Sicherheitshinweise vorliegen, und aktualisiert sie bei Bedarf.

## 4. Abnahmetest
1. App-Adresse öffnen → Zeile „☁ Anmelden“ ist sichtbar.
2. Anmelden mit einem Testkonto (ggf. Zustimmung bestätigen) → es erscheint der Name.
3. Plan anlegen, Bauvorhaben benennen, „☁ In Cloud speichern“ → Meldung „In Cloud gespeichert ✓“; im OneDrive bzw. in der Bibliothek liegt im Ordner `Verlegeplaene` die Datei `<Bauvorhaben>.vplan.json`.
4. Auf einem zweiten Gerät anmelden → „Cloud-Pläne“ → Plan öffnen.
5. Auf beiden Geräten ändern und speichern → die App warnt vor dem Überschreiben fremder Änderungen.
6. Variante A2: Zweite Person mit Bibliotheksrecht sieht den Plan; Person ohne Recht erhält eine Fehlermeldung.

## 5. Sicherheits- und Datenschutzhinweise
- Anmeldung nach OAuth 2.0 mit Autorisierungscode und PKCE; kein Client Secret im Code.
- Die Anmelde-Tokens liegen im Browser-Speicher des jeweiligen Geräts (localStorage). Auf gemeinsam genutzten Geräten muss man sich abmelden („Abmelden“-Knopf).
- Pläne werden als JSON-Dateien im Klartext in OneDrive/SharePoint abgelegt und unterliegen dort den üblichen Berechtigungen, Aufbewahrungs- und Löschrichtlinien.
- Die App-Dateien (inkl. Firmenlogo und Anschrift für den PDF-Kopf) sind bei einem öffentlichen GitHub-Repository öffentlich einsehbar. Wenn das nicht gewünscht ist: Betrieb auf dem Firmenserver.
- Die Pläne selbst werden nie auf GitHub abgelegt.

## 6. Was die Entwicklung braucht
Client-ID, Mandanten-ID, gewählte Variante (A1/A2), bei A2 Site-Adresse und Bibliotheksname sowie die endgültige App-Adresse.

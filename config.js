// Wird von der IT ausgefüllt (siehe IT-Einrichtung.md).
// Solange clientId leer ist, bleibt die Cloud-Funktion unsichtbar und die App arbeitet nur lokal.
window.VP_CLOUD = {
  clientId: '720429ef-95fc-4d3d-b097-04eee78cc078',          // Anwendungs-(Client-)ID der App-Registrierung in Microsoft Entra ID
  tenantId: '10af70ff-f844-4dfa-832f-b8f451b23e55',          // Verzeichnis-(Mandanten-)ID; leer = 'organizations' (jedes Firmenkonto)
  mode: 'sharepoint',      // 'onedrive' = eigenes OneDrive der angemeldeten Person
                         // 'sharepoint' = gemeinsame SharePoint-Bibliothek (alle Berechtigten sehen alle Pläne)
  folder: 'Verlegeplaene',               // Ordner, in dem die Pläne liegen (wird automatisch angelegt)
  site: { host: 'horstschulzbau.sharepoint.com', path: '/sites/Rohrleitungsbau-Verlegeplne', library: '' },
                         // nur bei mode 'sharepoint': host z. B. 'firma.sharepoint.com',
                         // path z. B. '/sites/Tiefbau', library z. B. 'Dokumente' (leer = Standardbibliothek)
  redirectUri: '',       // leer = Adresse der App (z. B. https://NAME.github.io/verlegeplan/)
  scopes: []             // leer = Standard (OneDrive: Files.ReadWrite, SharePoint: Files.ReadWrite.All)
};

# SV Ried Match Analysis

Komplett neue Streamlit-Basis für JWR / U18 / U16 / U15.

## Enthalten
- Direkte Team-Auswahl ohne Anmeldung
- Spiel anlegen
- Tor / Gegentor erfassen
- 3 Klickpunkte direkt am Spielfeld:
  1. Angriffsbeginn
  2. Assist / letzter Pass
  3. Abschluss
- Speicherung von X/Y-Koordinaten + automatisch abgeleiteter Zone
- Sichtbare 1–2–3-Markierung und Verbindungslinie direkt auf dem Spielfeld
- Dashboard mit Toren, Gegentoren, Abschlusszonen, Entstehung und Abschlusskarte
- Tore/Gegentore direkt aus der Ereignisliste löschen
- SQLite-Fallback für sofortigen lokalen Start
- Supabase-Unterstützung für produktiven Betrieb

## Lokal starten

```powershell
cd svried_goal_dashboard
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

Ohne `secrets.toml` startet die App direkt ohne Anmeldung und verwendet `data.db`.

## Supabase
1. Neues Supabase-Projekt anlegen.
2. `schema.sql` im SQL Editor ausführen.
3. `.streamlit/secrets.example.toml` nach `.streamlit/secrets.toml` kopieren.
4. `[supabase]` ausfüllen.
5. App neu starten.



## Datenmodell
Für räumliche Events werden immer Rohkoordinaten `x/y` (0..100) und zusätzlich die aktuell abgeleitete Zone gespeichert. Dadurch können später andere Zonen, Heatmaps und Vergleichsmodelle ergänzt werden, ohne Szenen neu erfassen zu müssen.

## Neu in v5
- Dunkles AKA-Dashboard im Look des bisherigen Boards
- Vier große Pitch-Ansichten im Stil des alten Dashboards
- Prozentwerte nach Abschlusszonen direkt im Pitch
- Weiterhin direkte Löschfunktion für Ereignisse

## Neu in v6
- Hauptansicht noch näher am bisherigen AKA Teams Dashboard
- Zwei dominante Pitch-Charts direkt nebeneinander
- kompakter FORZA-RIED/Header-Stil
- feinere Pitch-Geometrie und gestrichelte Hilfslinien
- Prozent-Badges im Stil des bisherigen Dashboards
- weitere Auswertungen in einen aufklappbaren Bereich verschoben
- Ereignis-Löschung bleibt direkt im Dashboard

## Neu in v7
- Dashboard-Grafik auf leeres Spielfeld reduziert
- alle Assist-/Hilfslinien entfernt
- 16m-Halbkreis korrigiert
- Punkte und Prozent-Badges bleiben erhalten

## Neu in v8
- Zonen im Dashboard wieder eingezeichnet
- keine diagonalen Assist-/Hilfslinien
- sauberes Spielfeld mit Zonenstruktur wie zuvor
- 16m-Halbkreis bleibt korrigiert

## Neu in v9
- rechter Hauptchart zeigt jetzt das ausgewählte Team statt alle Teams
- Beispiel U15: links `U15 - Eigene Tore`, rechts `U15 - Gegentore`

## Neu in v10
- Team-Auswahl und weitere Selectboxen im Dark-Theme wieder klar lesbar
- `Weitere Auswertungen` zeigen nur noch das aktuell ausgewählte Team
- keine `Alle Teams`-Beschriftung mehr in den Zusatzcharts

## Neu in v11
- gespeicherte Tore/Gegentore können wieder geöffnet werden
- komplette Szene mit Angriffsbeginn, Assist und Abschluss sichtbar
- Punkt 1, 2 oder 3 gezielt neu auf dem Spielfeld setzen
- Minute, Spieler, Spielphase, Entstehung, Video und Kommentar bearbeitbar
- Änderungen aktualisieren den bestehenden Datensatz
- Originalpunkte können vor dem Speichern wiederhergestellt werden

## Neu in v12
- Gegentore werden im Dashboard jetzt korrekt gespiegelt dargestellt
- Punkte aus dem Szenen-Editor und die Gegentor-Grafik stimmen jetzt überein
- Prozent-Badges für Gegentore werden ebenfalls auf Basis der gespiegelten Darstellung berechnet

## Neu in v13
- Lesbarkeit von Menüs, Überschriften und Formularen im Dark Theme deutlich verbessert
- Eingabefelder, Selectboxen, Textareas und Segment-Auswahl kontrastreicher
- Labels, kleine Überschriften und Expander-Titel klarer lesbar
- deaktivierte Buttons besser erkennbar

## Neu in v14
- Team-Auswahl im Sidebar deutlich lesbarer
- Ereignis-Auswahl `Tor / Gegentor` kontrastreicher und besser lesbar
- Bewerb beim Spiel anlegen jetzt als Auswahl `Testspiel` oder `Punktspiel`
- Spielauswahl zeigt nun zusätzlich den Bewerb an

## Neu in v15
- Sidebar-Team-Auswahl jetzt explizit dunkel mit weißer Schrift
- Dropdown-Pfeil und Label ebenfalls kontrastreich
- frühere CSS-Regeln werden durch spezifischere Sidebar-Regeln überschrieben

## Neu in v16
- Team-Auswahl nicht mehr als Dropdown
- U15 / U16 / U18 / JWR direkt als gut lesbare Auswahl in der Sidebar
- unabhängig vom Streamlit-Selectbox-Theme

## Neu in v17
- vor der Erfassung auswählbar: 1, 2 oder 3 Klicks
- 1 Klick = nur Abschluss
- 2 Klicks = Assist + Abschluss
- 3 Klicks = Angriffsbeginn + Assist + Abschluss
- Auswahl wird pro Ereignis gespeichert
- im Szenen-Editor kann die Anzahl der Punkte später geändert werden
- nicht benötigte Punkte werden beim Speichern geleert

## Neu in v18
- `Punkte der Szene` nicht mehr als Dropdown
- direkte Auswahl `1 Klick / 2 Klicks / 3 Klicks`
- gilt sowohl bei Erfassung als auch im Szenen-Editor
- damit unabhängig vom problematischen Selectbox-Theme

## Neu in v19
- zusätzlicher Bereich `Gesamtüberblick alle Teams`
- gemeinsame Pitch-Grafik für alle eigenen Tore
- gemeinsame Pitch-Grafik für alle Gegentore
- kompakte Teamübersicht U15 / U16 / U18 / JWR mit Tore, Gegentore und Differenz

## Neu in v20
- angelegte Spiele können wieder gelöscht werden
- Sicherheitsabfrage vor dem Löschen
- vorhandene Tore/Gegentore des Spiels werden mitgelöscht
- Anzahl Tore/Gegentore wird in der Spielverwaltung angezeigt

## Neu in v21
- eigener Sidebar-Menüpunkt `Gesamt Dashboard`
- Gesamtübersicht U15 / U16 / U18 / JWR mit Tore und Gegentore
- neues Merkmal `Torabschluss`: One Touch / Two Touch / >2 Touches
- Merkmal wird gespeichert und ist im Szenen-Editor bearbeitbar
- Touch-Auswertung im Team-Dashboard
- Touch-Auswertung im Gesamt-Dashboard über alle Mannschaften
- Mannschaftsvergleich mit One Touch / Two Touch / >2 Touches

## Neu in v22
- Assists als eigene Anzeige ergänzt
- im Team-Dashboard eigener Assist-Bereich
- im Gesamt-Dashboard eigener Assist-Bereich über alle Mannschaften
- in den Tore-Dashboards werden Abschluss und Assist mit transparenten Linien verbunden
- zusätzliche kleine Assist-Punkte in den Tore-Dashboards
- Gesamtvergleich jetzt auch mit Assist-Anzahl pro Mannschaft

## Neu in v23
- Koordinaten werden jetzt auf das tatsächliche Spielfeld statt auf den Bildrand normiert
- bestehende SQLite-Daten werden einmalig auf das neue Koordinatensystem umgerechnet
- Standardsituation vor der Punkterfassung auswählbar
- Eckball links/rechts setzt den Assistpunkt automatisch exakt auf den Eckpunkt
- Elfmeter setzt den Abschluss automatisch exakt auf den Elfmeterpunkt
- Standardsituation kann im Szenen-Editor später geändert werden

## Neu in v24
- aufgeklappte Dropdown-Menüs vollständig lesbar
- alle nicht ausgewählten Optionen: dunkler Hintergrund + weiße Schrift
- Hover-Zustand kontrastreicher
- ausgewählte Option im Dropdown grün hervorgehoben
- Fix gilt zentral für Bewerb, Spielphase, Entstehung und weitere Selectboxen

## Neu in v25
- Touch-Anteil-Grafik deutlich kompakter
- Assist-Pitch kleiner und zentriert
- gilt für Team-Dashboard und Gesamt-Dashboard

## Neu in v26
- Assist-Auswertung getrennt für eigene Tore und Gegentore
- Team-Dashboard zeigt beide Assist-Pitches nebeneinander
- Gesamt-Dashboard zeigt beide Assist-Pitches nebeneinander
- auch im Gegentore-Dashboard werden Assist und Abschluss transparent verbunden
- Gesamtvergleich enthält Assists eigene Tore und Assists Gegentore separat

## Neu in v27
- `Assists eigene Tore` jetzt in derselben Pitch-Größe wie `Eigene Tore`
- gilt im Team-Dashboard und im Gesamt-Dashboard
- `Assists Gegentore` bleibt weiterhin kompakter

## Neu in v28
- Sidebar oben nur noch mit SV Ried Logo
- Text `⚽ SV Ried` entfernt
- `AKA Goal Dashboard` in der Sidebar entfernt
- Hinweis `U15 · U16 · U18 · JWR gemeinsam` im Gesamt-Dashboard entfernt

## Neu in v29
- Sidebar-Logo kleiner dargestellt
- Logo weiter nach links oben verschoben

## Neu in v30
- `Assists eigene Tore` wird jetzt auch ohne erfasste Assists angezeigt
- mit Pitch-Grafik, Anzahl Assists = 0 und Hinweistext
- gilt im Team-Dashboard und im Gesamt-Dashboard

## Neu in v31
- `Assists Gegentore` wird auch bei 0 Assists immer mit Pitch angezeigt
- Assist-Grafiken eigene Tore und Gegentore sind gleich groß
- gilt im Team-Dashboard und im Gesamt-Dashboard


## v32 – Supabase / PostgreSQL produktiv

### 1. Supabase-Projekt anlegen
Neues Supabase-Projekt erstellen und anschließend im **SQL Editor** den kompletten Inhalt von `schema.sql` ausführen.

### 2. Lokale Secrets
`.streamlit/secrets.example.toml` nach `.streamlit/secrets.toml` kopieren und URL + serverseitigen Supabase Secret/Service-Role-Key eintragen.

### 3. Verbindung testen
```powershell
python check_supabase.py
```

### 4. Bestehende SQLite-Daten übertragen
Wenn die aktuelle `data.db` übernommen werden soll:
```powershell
python migrate_sqlite_to_supabase.py
```

### 5. Lokal starten
```powershell
python -m streamlit run app.py
```
Sobald `[supabase]` in `secrets.toml` vorhanden ist, verwendet die App automatisch PostgreSQL statt SQLite.

### 6. Streamlit Cloud
Unter **App -> Settings -> Secrets** eintragen:
```toml
[supabase]
url = "https://YOUR_PROJECT.supabase.co"
key = "YOUR_SERVER_SIDE_SECRET_KEY"
```

`data.db` und `.streamlit/secrets.toml` sind über `.gitignore` ausgeschlossen.

## Neu in v33
- Spielfeld bei Tor-/Gegentor-Erfassung vertikal
- Angriffsrichtung in Erfassung und Szenen-Editor von unten nach oben
- internes 0–100-Koordinatensystem bleibt unverändert, bestehende Daten bleiben kompatibel
- neue Torzeit-Auswertung pro Mannschaft und im Gesamt-Dashboard
- Zeitfenster: 1–25, 26–45, 46–75, 76–90
- Tore und Gegentore werden je Zeitfenster getrennt dargestellt

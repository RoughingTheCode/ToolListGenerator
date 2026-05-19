
# Tool List Generator

> Erstellt eine Werkzeugliste anhand einer von der Maschine geladenen oder einer lokal geöffneten Datei. Optional können die Daten mit den in der Maschine vorhandenen Werkzeugen abgeglichen werden. Die Werte werden anschließend direkt in eine automatisch generierte Excel-Werkzeugliste eingetragen.

<img width="688" height="767" alt="TLG_Programm_Bild" src="https://github.com/user-attachments/assets/712e2c6d-ffa0-44dd-917e-fe47cb6e291e" />

<img width="681" height="762" alt="TLG_M_Einstellungen_Bild" src="https://github.com/user-attachments/assets/b9d90e7a-c30d-4d3c-a151-85672232a9f0" />

<img width="676" height="756" alt="TLG_W_Einstellungen_Bild" src="https://github.com/user-attachments/assets/846c87e8-b612-4edc-af4b-a8d9907af16b" />


## Inhaltsverzeichnis

- [Hinweis](#-hinweis)
- [Voraussetzungen](#-voraussetzungen)
- [Installation & Einrichtung](#-installation--einrichtung)
- [Verwendung](#-verwendung)
  - [Suchkriterien](#-suchkriterien)
- [Mitwirken](#-mitwirken)
- [Lizenz](#-lizenz)

---

## ⚠️ Hinweis

> **Wichtig:** Dies ist ein reines Hobbyprojekt und befindet sich noch in der Entwicklung. Die Auswertung von NC-Code und der automatisierte Maschinendownload sind komplexe Prozesse. Da NC-Dialekte stark variieren können, besteht immer das Risiko von Fehlinterpretationen.

**Die Nutzung erfolgt auf eigene Gefahr. Der Entwickler übernimmt keinerlei Haftung.**

* **Unterstützte Steuerungen:** Bisher ist das Programm nur für Maschinen mit Heidenhain-Steuerungen entwickelt. Eine Erweiterung ist für die Zukunft geplant.
* **Einschränkungen:** Das Programm liest aktuell nur Werkzeugnummern, keine Strings (genaue Suchkriterien werden unter [Verwendung](#-verwendung) erklärt).
* **Datenschutz:** Aus Datenschutzgründen kann die Benutzung des Programms leider nicht in einem Video veranschaulicht werden.

---

## 💻 Voraussetzungen

Vor der Nutzung sollten folgende Programme auf dem PC installiert sein:
- [Microsoft Excel](https://www.microsoft.com/de-de/microsoft-365/excel)
- [Heidenhain TNCremo](https://www.heidenhain.de/service/downloads/software)

*Tipp:* Um das Programm im vollen Umfang nutzen zu können, sollten die Maschinen an das Netzwerk angeschlossen sein. Es gibt jedoch auch eine Funktion (`Programm öffnen`), mit der man lokal eine Datei wählen kann. In diesem Fall muss bei den Maschinen-Einstellungen *keine* Maschine ausgewählt sein.

---

## 🚀 Installation & Einrichtung

1. **Programm starten:** Den Ordner mit der `TLG.exe` auf den PC kopieren. Die Datei ist direkt ausführbar (es wird keine Installation benötigt). Bei Bedarf kann eine Verknüpfung auf dem Desktop oder in der Taskleiste angelegt werden.
2. **Einstellungen öffnen:** Die `.exe` starten und direkt in den Reiter **Einstellungen** wechseln. Dort sind standardmäßig Beispieleinstellungen hinterlegt.
3. **TNCcmd verknüpfen:** Ganz oben auf *Datei wählen* klicken und die `TNCcmd.exe` im installierten TNCremo-Verzeichnis suchen und auswählen.
4. **Maschine anlegen:**
   - Einen **Namen** für die Maschine vergeben.
   - Die **IP-Adresse** der Maschine eingeben. *(Falls keine Verbindung aufgebaut wird, teste es direkt in TNCremo. Gegebenenfalls muss an der Maschine erst freigegeben werden, dass sie ins Netzwerk darf).*
   - Den **Hauptordner-Pfad** auswählen (von dort aus wird später auch in den Unterordnern gesucht).
   - Pfad zur **Werkzeugtabelle** eintragen (falls kein Werkzeugwechsler vorhanden ist, das Feld einfach leer lassen).
   - Pfad zur **Platztabelle** eintragen (auch hier: falls kein Werkzeugwechsler vorhanden, leer lassen).
5. **Erweiterte Ordner-Einstellungen:**
   - Im Reiter **Unterordner** können weitere Ordner kommagetrennt eingetragen werden. Diese lassen sich später über ein Dropdown-Menü auswählen.
   - **Ordner-Struktur:** Hier `Ja` oder `Nein` wählen. Dies ist nützlich, wenn Programme nach Firmenkürzeln sortiert sind. 
     *Beispiel:* Firma Haribo, Kürzel `HAB`, Programm heißt `HAB_55789`. Das Programm liest anhand der definierten Zeichen das Kürzel aus und wählt automatisch den Ordner `HAB`. Hierzu muss man anschließend die **Anzahl und Position der Zeichen** festlegen (z. B. 1 bis 3 für `HAB`, oder 5 bis 7 für die Nummer).
6. **Werkzeuglisten-Einstellungen:**
   - Über den Button **WKZ-Listen Einstellungen** das Menü wechseln (mit *Maschinen Einstellungen* kommt man wieder zurück).
   - Im obersten Eingabefeld können **Werkzeuge ausgeschlossen** werden, die sich standardmäßig immer in der Maschine befinden (kommagetrennt).
   - Bei **IKZ-Funktionen** die M-Funktionen für die Innenkühlung eintragen. Das Programm trägt dann in die Liste ein, ob es sich um ein `A` (Außengekühltes Werkzeug) oder ein `I` (Innengekühltes Werkzeug) handelt.
   - Unter **Spann-Notizen** einstellen, ob diese mit eingelesen werden sollen (werden über das Kommentarzeichen `;` gesucht).
   - Unter **Werkzeuge sortieren** wählen, ob die Liste nach Nummern sortiert werden soll.
7. **Speichern:** Abschließend auf **Speichern** drücken. Die Maschine erscheint nun oben im Dropdown-Feld. *(Tipp: Wenn du die Beispielmaschine auswählst, kannst du sie über den Löschen-Button direkt entfernen. Wenn du später Einstellungen überschreiben willst, muss der Name 1:1 übereinstimmen. Ändert sich auch nur ein Zeichen, wird eine neue Maschine gespeichert).*

---

## 🛠 Verwendung

1. Im Dropdown-Feld **Maschine** die gewünschte Maschine auswählen.
2. Im Feld **Ordner** kann bei Bedarf ein anderer Unterordner gewählt werden. Steht die Auswahl auf *Bezugsordner*, wird vom hinterlegten Hauptpfad heruntergeladen.
3. Bei den **Optionen** können die Reiter *Spann-Notizen* und *Werkzeug sortieren* für den aktuellen Lauf nochmals temporär angepasst werden.
4. Im Eingabefeld **Programme** die gewünschten Programmnummern kommagetrennt eintragen.
5. Auf **WKZ-Liste generieren** klicken. Es ploppen nun kurz schwarze Fenster auf (für den Download der Programme, der Werkzeugliste und der Platztabelle).
   - *Hinweis:* Die geladenen Programme landen im Ordner `Programm`, die Platztabelle in `Platz_Tabelle` und die Maschinen-Werkzeugliste in `Werkzeuglisten`. Diese Ordner werden bei jeder neuen Ausführung zu Beginn geleert. Die fertige Excel-Liste landet im Ausgabeordner.
6. **Ordner-Struktur in Aktion:** Wenn die Ordner-Struktur aktiv ist, ändert sich beim Eintragen der Programmnummer automatisch der Text bei **Aktiver_Pfad**. So siehst du direkt, welcher Zielordner ausgewählt wurde. Die manuelle Auswahl eines Unterordners überschreibt/deaktiviert diese Automatik.
7. **Lokale Dateien:** Über den Button **Programm öffnen** kann ein lokales Programm vom PC ausgewählt werden. Die Werkzeuge der ausgewählten Maschine werden auch hier abgeglichen. Wählt man bei Maschine *keine* (oder hat keine Tabellen hinterlegt), funktioniert das Programm komplett lokal ohne Maschinenanbindung.

> *Info zur Performance:* Der Ordner `_internal` entsteht beim Erstellen der `.exe`. Dies sorgt dafür, dass das Programm wesentlich schneller läuft, als wenn alles in eine einzige, große `.exe` gepackt worden wäre. Die `.ico`-Dateien für die Programm-Icons liegen im Ordner `Bilder`. Deine Einstellungen werden im Ordner `Einstellungen` als `.json`-Datei gespeichert – diese kannst du bei Bedarf einfach auf einen anderen PC kopieren, um dir die erneute Einrichtung zu sparen.

### 🔍 Suchkriterien

Wie genau liest das Programm den NC-Code aus?

- **Gliederungen:** Das Programm liest zuerst alle Gliederungen aus. Informationen wie Kunde, Zeichnungsnummer oder Revision (Index, Version) werden erkannt und in die Werkzeugliste eingetragen.
- **CAM-Erkennung:** Es gibt eine einfache Überprüfung, ob es sich um ein CAM-Programm handelt (Erkennung über die Anzahl an `-` Zeichen am Programmanfang, ca. 37 Stück). Ist dies der Fall, wird der oberste Kommentar als Spann-Notiz gewertet. Die Begriffe "Erzeugt am" und "Datei" dienen als Suchkriterien und werden (sofern gefunden) eingetragen.
- **Manuelle Programme:** Handelt es sich nicht um ein CAM-Programm und ist "Mit Spann-Notizen" aktiviert, werden alle Kommentare (mit `;` am Anfang) in das Feld *Beschreibung/Aufspannung* der Werkzeugliste eingetragen, bis das erste Werkzeug gefunden wird.
- **Werkzeug-Erkennung:** Das Werkzeug wird immer aus der Gliederung ausgelesen. Das Suchkriterium lautet `* - T(Werkzeugnummer)`. Es wird strikt geprüft, ob es sich um eine Zahl handelt, um versehentliche Treffer wie `* - Teile nicht mit Magnet heben` auszuschließen (daher funktionieren aktuell keine Werkzeug-Strings). Alles, was in der Zeile hinter der Werkzeugnummer steht, landet in Excel in der *Beschreibung*.
- **Kommentare nach Werkzeugaufruf:** Befindet sich unmittelbar nach dem Werkzeugaufruf ein Kommentar (`;`), wird dieser der Beschreibung beigefügt.
- **Kühlung (IKZ):** Das Programm sucht nach den eingestellten M-Befehlen. Wird einer gefunden, wird ein `I` (Innenkühlung) eingetragen, andernfalls ein `A` (Außenkühlung).
- **Alte Werkzeuge ignorieren:** Wenn im Programm unten eine Sicherung mit altem Werkzeug angefügt ist, kann diese über eine Gliederung mit `* - SICHERUNG`, `* - Sicherung` oder `* - sicherung` markiert werden. Das Programm hört an dieser Stelle auf, weiterzulesen.

---

## 🤝 Mitwirken

Da ich nur begrenzte Testdaten zur Verfügung habe, freue ich mich sehr, wenn jemand den Code für sich anpasst, testet oder das Projekt erweitert. 
Gerne könnt ihr mich bei Fragen oder Anpassungsbedarf auch direkt anschreiben!

---

## ☕ Unterstütze dieses Projekt

Dieses Projekt wird in meiner Freizeit entwickelt. Wenn dir der **Tool List Generator** den Arbeitsalltag erleichtert und du meine Arbeit unterstützen möchtest, freue ich mich riesig über einen virtuellen Kaffee!

[![Ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/DEIN_KOFI_NAME)

---

## 📄 Lizenz

MIT License

Copyright (c) 2026 RoughingTheCode

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

\## Tool List Generator

> Erstellt eine Werkzeugliste anhand einer von der Maschine Geladenen Datei oder einer Lokal geöffneten Datei und gleicht wenn gewollt diese mit den sich in der Maschine befindenden Werkzeugen ab und trägt direkt werte in die in Excel erstellte Werkzeugliste ein.







\## Inhaltsverzeichnis



- \[Hinweis](#-hinweis)

- \[Installation](#-installation)

- \[Verwendung](#-verwendung)

- \[Mitwirken](#-mitwirken)

- \[Lizenz](#-lizenz)







\## Hinweis



> \*\*Wichtig\*\* dies ist ein reines Hobbyprojekt und befindet sich noch in der Entwicklung. Die Auswertung von NC-Code und der automatisierte Maschinendownload sind komplexe Prozesse.

Da NC-Dialekte Stark variieren können besteht immer das Risiko von Fehlinterpretation.



> \*\*Die Nutzung erfolgt auf eigene Gefahr. \*\*Der Entwickler übernimmt keinerlei Haftung.



> \*\* Bisher ist das Programm nur für Maschinen mit Heidenhain Steuerungen Entwickelt dies soll mit der Zeit noch erweitert werden.



- Das Programm liest aktuell nur Werkzeug Nummern keine Strings genaue suchkriterien erkläre ich bei Verwendung.

- Aufgrund von Datenschutzgründen kann ich leider das Benutzen des Programms nicht zur veranschaulichung in einem Video zeigen.





\## Installation



> \*\*Vor der Installation sollten folgende Programme auf dem PC vorinstalliert sein:



- Excel(https://www.microsoft.com/de-de/microsoft-365/excel)

- TNCremo(https://www.heidenhain.de/service/downloads/software)



- Um das Programm im vollen umfang nutzen zu können sollte man die Maschinen im Netzwerk angeschlossen haben (es gibt auch eine funktion über Programm öffnen mit der man lokal eine datei wählen kann in den Maschinen muss dann _keine_ gewählt sein).







- Einfach Ordner mit der TLG.exe auf den PC ziehen eine falls eine Version auf dem Desktop oder unten in der Leiste erwünscht sind eine verknüpfung reinziehen die exe sollte direkt ausführbar sein.

- Die exe starten und direkt in die Einstellungen wechseln dort befinden sich Standardmäßig Beispiel Einstellungen.

- Als allererstes ganz oben Datei wählen anklicken und wie Beschriebene TNCcmd.exe im Installierten TNCremo Verzeichnis suchen und auswählen.

- Danach einen Namen für die Maschine vergeben.

- Die IP Adresse der Maschine eingeben (Falls diese nicht geht in TNCremo mal probieren sollte es dort auch nicht gehen muss man eventuell erst auf der Maschine freigeben das diese ins Netzwerk darf).

- Den Haupt Ordner Pfad auswählen (von dort aus wird später auch in den unterordnern gesucht).

- Pfad zur Werkzeugtabelle eintragen(falls kein Werkzeug wechsler vorhanden Feld einfach leer lassen).

- Pfad zur PLatztabelle eintragen (gleiches spiel auch hier falls kein Werkzeugwechsler vorhanden leer lassen).

- Im reiter Unterordner weitere ordner kommagetrennt eintragen diese kann man später über ein Dropdown Menu auswählen.

- Ordner Struktur ja oder nein wählen Falls man zum beisspiel die Programme nach Firmen sortiert hat und jede Firma ein gleich langes kürzel hat kann man hier Ja eintragen eine bestimmte anzahl an zeichen wird hierdurch ausgewählt als Ordner name das programm muss aktuell in dieser version jedoch auch diese zeichen haben beispiel Firma Haribo kürzel HAB pgm HAB_55789 wird der Ordner HAB ausgewählt das könnte man gegebenenfalls falls gebraucht auch abändern das HAB_ einfach Ordner HAB ist und erst die nummer oder die zeichenreihe danach das programm anwählt.

- Anschließend die anzahl und die Position der Zeichen wählen 1 bis 3 wäre Bei HAB_88989 HAB, Zeichen 5-7 wäre 889

- Nun oben auf den Button WKZ-Listen Einstellungen klicken um das menü zu wechseln mit dem Button Maschinen Einstellungen kommt man wieder zurück.

- Hier können nun im oberten Eingabefeld Werkzeuge ausgeschlossen werden die immer in der Maschine sind diese werden wieder mit Komma getrennt.

- Bei IKZ-Funktionen die M-Funktionen für IKZ eintragen in die WKZ-Liste wird Anschließend eingetragen ob es sich um ein A=Außengekühltes Werkzeug oder I=Innengekühltes Werkzeug handelt.

- Beim reiter Default Mit Spann-Notizen wird eigestellt ob man Spann Notizen mit einlesen will oder nicht (diese werden über kommentar ; gesucht).

- Beim reiter Werkzeuge Sortieren kann man Auswählen ob diese nach nummern sortiert werrden sollen

- Anschließend auf Speichern drücken nun erscheint die Maschine oben im Dropdown Feld wenn man die Beispiel Maschine auswählt kann man diese über den Löschen Button gleich entfernen.

- Will man später Einstellungen überschrieben muss man den Namen eins zu eins stehen lassen ändert man ein zeichen Speichert man eine Neue Maschine





## Verwendung



- Wenn alles eingerichtet ist kann man nun beim Dropdown Feld mit der Überschrift Maschine die gewünschte Maschine auswählen.

- Im Dropdown Feld Unterordner kann man wenn gewünscht einen anderen Ordner in dem sich das Programm befindet auswählen wenn _Bezugsordner_ ausgewählt ist wird das Programm vom Bezugsordner Pfad aus heruntergeladen.

- Bei den Optionen können die reiter Spann-Notizen und Werkzeug sortieren nochmal geändert werden wenn gewünscht.

- Im eingabe Feld Programme können nun mehrere gewünschte Programme die mit einem komma getrennt werden müseen Eingetragen werden über den Button WKZ-Liste generieren wird nun der Prozess gestartet es plopen Schwarze Fenster auf für jedes Programm das geladen wird und für das Laden der Werkzeug_Liste und der PLatztabelle.

- Die Programme werden im Programm ordner gespeichert die Platz Tabelle im Platz_Tabelle Ordner und die Werzeugliste im Werkzeuglisten Ordner diese werden bei jedem benutzen zu beginn abgelöscht die fertige Excel Werzeugliste landet im Ausgabeordner. Im Ordner Bilder sind Ico dateien für die Programm Icons im ordner Einstellungen ist eine Json datei mit den einstellungen diese kann einfach auf ein zweites gerät kopiert werden damit spart man sich die arbeit diese neu einzutragen.

- Der Ordner _internal entsteht beim erstellen der exe diese läuft so wesentlich schneller als wenn man alles in der einen exe erstellen lässt.

- Beim eintragen der Programm nummer merkt man nun falls der reiter Ordner Struktur aktiv ist das die Zeile Aktiver_Pfad sich verädert wenn man das Programm einträgt damit wird angezeigt welcher ordner nun als Zielordner der Datei ausgewählt ist. Falls man einen Unterornder auswählt erscheint dieser dort auch die einstellung unterordner überschreibt die Ordner-Struktur schaltet diese sozusagen aus.

- Über den Button Programm öffnen kann man ein Lokales Programm auswählen über das eine Werkzeugliste generiert werden soll die Werkzeuge der ausgewählten Maschine werden auch hier abgeglichen ausser man wählt bei Maschine _keine_ oder hat keine Werkzeugliste und Platztabelle eingetragen nun funktioniert das Programm einfach lokal ohne Maschinen Anbindung.





** Suchkriterien



- Das Programm liest zuerst alle Gliederungen aus hier haben wir infos wie Kunde Zeichnungsnummer Revision(Index, Version) stehen und trägt diese immer in die Werzeugliste ein. Es gibt im Python Script eine ausnahme hier wird überprüft ob es ein CAM programm ist allerding sehr primitiv über die anzahl an "-" (wenn ich richtig gezählt habe 37) Zeichen am anfang des Programmes dann wird der Kommentar oben als Spannotiz angesehen und Erzeugt am und Datei dienen als Suchkriterien in Sätzen und werden wenn gefunden immer eingetragen.

- Sollte es sich nicht um ein CAM programm handeln trägt er wenn Mit Spann-Notizen an ist alle kommentare(; am anfang) und trägt diese ebenfalls unter die Gliederungen in das Feld Beschreibung/Aufspannung in der Werkzeugliste ein dies macht er solange bis er ein Werkzeug findet.

- Das Werzeug liest das programm aufgrund zusätzlich benötigter Infos wie werkzeug Name immer aus der Gliederung aus das Suchkriterium ist hier * - T(Werkzeugnummer) es wird im Programm geprüft ob es sich dabei um eine nummer handelt um dinge wie "* - Teile nicht mit Magnet heben" aus dem suchkriterium für Werkzeuge auszuschließen daher können Werkzeug Strings aktuell nicht verwendet werden. Alles hinter der Werkzeugnummer wird in Excel in der Beschreibung eingetragen.

- Sollte sich unmittelbar nach dem Werkzeug ein kommentar (;) befinden wird dieser ebenfalls der Beschreibung beigefügt.

- Das Programm sucht anschließend nach den Eingestellten M-Befehlen für die IKZ und wenn es einen findet wird ein I für Innenkühlung eigetragen sollte es nichts finden steht dort immer ein A für Außenkühlung.

- Falls man dinge ändert und im Programm eine Sicherung mit altem Werkzeug unten anfügt kann man diese über eine Gliederung mit:

"* - SICHERUNG" , "* - Sicherung" oder "* - sicherung" makieren das Programm hört an diesen stellen auf weiter zu lesen.







## Mitwirken

- Da ich nur begrenzte test Daten zu verfügung habe freue ich mich wenn jemand sich den Code auf sich anpasst oder ihn ergänzt.

- Gerne kann man mich auch bei anpassungsbedarf anschreiben.





## Lizenz

-

Hier ist deine überarbeitete README.md. Ich habe die Rechtschreib- und Grammatikfehler korrigiert, extrem lange Schachtelsätze in verständliche Schritte aufgeteilt und das Layout durch Code-Hervorhebungen (Beispiel) und übersichtliche Listen optisch aufgewertet.

Außerdem habe ich am Ende, wie gewünscht, die MIT-Lizenz mit deinem GitHub-Namen und dem Jahr 2026 eingefügt.

Du kannst den Code hier einfach kopieren und in dein GitHub-Repository einfügen:

Markdown
# Tool List Generator

> Erstellt eine Werkzeugliste anhand einer von der Maschine geladenen oder einer lokal geöffneten Datei. Optional können die Daten mit den in der Maschine vorhandenen Werkzeugen abgeglichen werden. Die Werte werden anschließend direkt in eine automatisch generierte Excel-Werkzeugliste eingetragen.

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

import xml.etree.ElementTree as ET
import json
import ollama
import pymupdf
import os
import openpyxl
import numpy
import re
from pydantic import BaseModel, Field, ValidationError
from decimal import Decimal
from copy import deepcopy
import pdf_report
from pathlib import Path

class AIPositions(BaseModel):
    positionsnummer: str | None = Field(
        default=None, 
        description="Kennzeichnungsnummer der Position. Exakt wie im Soll-LV-Ausschnitt und stehen meist auf der Linken seite und über dem Einheitspreis."
    )
    einheitspreis: float | None = Field(
        default=None, 
        description="Der Einheitspreis dieser Position. Handschriftlich eingetragen neben der Beschriftung Einheitspreis"
    )
    gesamtpreis: float | None = Field(
        default=None, 
        description="Der Gesamtpreis dieser Position. Steht rechts neben dem Einheitspreis, wird auch Positionspreis genannt."
    )
    Unsicherheit: bool = Field(
        default=False, 
        description="setze es dieses Feld True falls du dir beim erkennen der Position unsicher bist oder sie nicht lesen kannst."
    )
    bieterkommentar: str | None = Field(
        default=None, 
        description="Feld für handschriftliche Kommentare die der Bieter zusätzlich zur Positionsbeschreibung hinterlassen hat. Falls handschriftlicher Kommentar vorhanden unbedingt einfügen."
    )

class positionList(BaseModel):
    positions: list[AIPositions] = Field(default_factory=list)
    angebots_preis: float | None = Field(
        default=None, 
        description="Der Angebotspreis steht nur am Schlussblatt und ist die Summe aller positionen ohne Steuern und Nachlass/Skonto"
    )
    AIkommentar: str = Field(
        default="", 
        description="Optionaler Kommentar den du bei Unsicherheiten und Auffälligkeiten hinterlassen kannst. z.b Position die durch Seitenumbruch ageschnitten ist"
    )

class FinalResult(BaseModel):
    positions: list[AIPositions] = Field(
        default_factory=list, 
        description="Liste aller extrahierten Einzelpositionen und deren Preisen"
    )
    angebots_preis: float | None = Field(
        default=None, 
        description="Der Angebotspreis steht nur am Schlussblatt und ist die Summe aller Gesamtpreise der positionen ohne Steuern und Nachlass/Skonto"
    )
    fehlende_positionen: list[str] = Field(
        default_factory=list, 
        description="Liste von Positionsnummern oder Bezeichnungen, die im Angebot fehlen aber im Soll_LV stehen."
    )
    doppelte_positionen: list[str] = Field(
        default_factory=list, 
        description="Liste von Positionsnummern, die fälschlicherweise mehrfach auftauchen."
    )
    Bericht: str = Field(
        default="", 
        description="Zusammenfassender Textbericht oder Analyse der Überprüfung."
    )
    
class Position:
    def __init__(self, positionsnummer, description, lvmenge, einheit, element = None):
        self.positionsnummer = positionsnummer
        self.description = description
        self.einheit = einheit
        self.lvmenge = lvmenge
        self.element = element
        self.einheitspreis = None
        self.gesamtpreis = None

PROMPT = """Das Bild ist ein Ausschnitt einer Seite aus einem Leistungsverzeichnis (Bauwesen). Der Bieter hat die Seite ausgedruckt, Preise und Anmerkungen von Hand eingetragen und eingescannt. Lies die Positionsnummern und die handschriftlichen Preise aus. Unter dieser Anweisung stehen die wahrscheinlichste Position und die Soll-LV-Positionen dieser Seite. Beides ist nur Orientierung und kann falsch sein. Trage ein, was du im Bild siehst, nicht was du erwartest.

AUFBAU EINER POSITION (von oben nach unten):
1. Links oben steht die gedruckte Positionsnummer, daneben das Stichwort. Danach folgt der oft lange Positionstext.
2. Am Ende stehen die handschriftlichen Preise. Der Einheitspreis steht auf der Zeile mit der Beschriftung "Einheitspreis". Darüber stehen manchmal Lohn und Sonstiges, diese Werte sind nicht gesucht. Auf derselben Zeile ganz rechts, nach "EUR", steht der Gesamtpreis.
3. Darunter beginnt die nächste Position.

AUSSCHNITT UND SEITENUMBRUCH:
- Ein Ausschnitt beginnt meist an einer Positionsnummer und endet vor der nächsten.
- Der erste Ausschnitt einer Seite beginnt am Seitenanfang, meist mit dem Seitenkopf. Danach folgt oft noch das Ende der letzten Position der Vorseite (Positionstext oder Preise, ohne Nummer).
- Der letzte Ausschnitt einer Seite endet am Seitenende. Dort kann eine Position abgeschnitten sein.
- Bei einem Seitenumbruch kann eine Position geteilt sein: Auf einer Seite stehen Nummer und Positionstext, auf der nächsten erst die Preise. Trage jeweils nur ein, was im Bild zu sehen ist.

POSITIONSNUMMERN:
- Trage jede Positionsnummer ein, die du siehst, auch wenn Preise fehlen oder die Position abgeschnitten ist. Siehst du etwas, das wie eine Positionsnummer aussieht, trage es ein, auch wenn du nicht sicher bist, ob es eine echte Position ist. Was keine echte Position ist, sortiert der Code danach aus.
- Schreibe die Nummer so ab, wie sie gedruckt steht, zum Beispiel 07.15.03.12 oder 07.15.03.12.A. Erfinde nie eine Nummer.
- Seitenkopf, Fußzeile und Überträge sind keine Positionen. Trage sie nicht ein, auch wenn Beträge dabei stehen.

PREISE:
- Nur abschreiben, nichts nachrechnen, nichts schätzen. Preis als normale Zahl wie in JSON, ohne Tausendertrennzeichen: 12,50 wird 12.5 und 1.234,50 wird 1234.5.
- Jeder Preis zählt einzeln. Was lesbar dasteht, wird eingetragen. Ist ein Preis leer oder durchgestrichen, ist nur dieser Preis None, der andere Preis bleibt davon unberührt.
- "0", "inkl." und "enthalten" ergeben 0.
- Wurde ein Preis korrigiert oder überschrieben, gilt der neue Wert. Setze Unsicherheit true und vermerke die Korrektur im bieterkommentar.
- Ist ein Preis nicht sicher lesbar, schreibe None und setze Unsicherheit true, statt zu raten.

WAS ZURÜCKGEGEBEN WIRD:
positions enthält einen Eintrag pro sichtbarer Positionsnummer. Sind mehrere Nummern sichtbar, gib mehrere Einträge zurück.
- Nummer und mindestens ein Preis sichtbar: positionsnummer eintragen, jeden lesbaren Preis eintragen, den Rest None, Unsicherheit false.
- Nummer und Preiszeile sichtbar, aber kein Preis eingetragen: positionsnummer eintragen, beide Preise None, Unsicherheit false.
- Nummer sichtbar, aber der Ausschnitt endet vor der Preiszeile: positionsnummer eintragen, beide Preise None, Unsicherheit true.
- Nur die Preise einer Position sichtbar (Zeile "Einheitspreis"), ohne Nummer: positionsnummer None, Preise eintragen, Unsicherheit true.
- Weder Nummer noch Preise sichtbar: positions = [].
- Handschriftlicher Kommentar ohne Position: ein Eintrag, alle Felder None außer bieterkommentar.

BIETERKOMMENTAR:
Alle handgeschriebenen Wörter und Anmerkungen des Bieters (z. B. "siehe Pos. ...", "Alternative", "inkl."), wörtlich. Also nur handschriftliche Sätze und Worte die als zusatzinformationen vom Bieter ergänzt wurden.

UNSICHERHEIT UND AIkommentar:
Unsicherheit true, wenn Nummer oder Preis unlesbar, mehrdeutig oder unsicher erkannt sind, sonst false. AIkommentar nur ausfüllen, wenn es etwas zu erklären gibt (ein Satz), sonst leer.

ENDSUMME:
Zeigt der Ausschnitt die Aufzählung aller Zwischensummen mit einer Endsumme, trage die Endsumme vor Mehrwertsteuer, vor Nachlass und vor Skonto in angebots_preis ein. Sonst None.

BEISPIELE (die Werte sind erfunden, nur das Muster zählt):
Normale Position:
{"positions":[{"positionsnummer":"07.15.03.12.A","einheitspreis":84.5,"gesamtpreis":338.0,"Unsicherheit":false,"bieterkommentar":null}],"angebots_preis":null,"AIkommentar":""}
Ausschnitt endet vor der Preiszeile:
{"positions":[{"positionsnummer":"07.15.03.13","einheitspreis":null,"gesamtpreis":null,"Unsicherheit":true,"bieterkommentar":null}],"angebots_preis":null,"AIkommentar":"Preiszeile nicht sichtbar, Position am Ende des Ausschnitts abgeschnitten."}
Nur Preise ohne Nummer:
{"positions":[{"positionsnummer":null,"einheitspreis":12.3,"gesamtpreis":36.9,"Unsicherheit":true,"bieterkommentar":null}],"angebots_preis":null,"AIkommentar":"Preise ohne Positionsnummer am Anfang des Ausschnitts."}
Zwei Positionen, die erste mit "inkl.":
{"positions":[{"positionsnummer":"07.15.04.01","einheitspreis":0,"gesamtpreis":0,"Unsicherheit":false,"bieterkommentar":"inkl."},{"positionsnummer":"07.15.04.02","einheitspreis":19.9,"gesamtpreis":99.5,"Unsicherheit":false,"bieterkommentar":null}],"angebots_preis":null,"AIkommentar":""}
Nichts Relevantes:
{"positions":[],"angebots_preis":null,"AIkommentar":""}
"""

SYSTEM_PROMPT = """
1. ROLLE UND ZIEL
Du bist ein Prüf-Agent für die automatische Auslesung handschriftlich ausgefüllter Leistungsverzeichnisse (Bauwesen).
Ein Vision-Modell hat ein eingescanntes Angebot ausschnittweise analysiert und Positionsnummern, Einheitspreise und Gesamtpreise extrahiert. Dabei können Lesefehler entstehen.
Deine Aufgabe ist es, die Ergebnisse zu überprüfen, eindeutige Fehler zu korrigieren und ungelöste Probleme zu dokumentieren.
Der ursprüngliche Scan ist die maßgebliche Quelle. Lieber einen Eintrag unverändert melden als falsch korrigieren.
2. GRUNDSÄTZLICHE REGELN
•	Positionsnummern und Preise dürfen ausschließlich auf Grundlage eindeutiger visueller Erkenntnisse aus rescan_page geändert werden. get_position dient nur zur Abfrage vorhandener Daten und darf niemals allein als Grundlage einer Korrektur dienen. Änderungen werden ausschließlich mit update_position durchgeführt.Rechnerische Auffälligkeiten beweisen keinen Lesefehler.
•	Erfinde oder berechne keine fehlenden Werte.
•	Ändere nur etwas, wenn die visuelle Überprüfung eindeutig ist.
•	Unbekannte Nummern und fehlende Positionen dienen als Hinweise zur Aufklärung anderer Probleme.
•	Die Positionsnummer "1" ist ein bekannter Formatfehler und wird ignoriert.
3. ABLAUF DER PRÜFUNG
Phase 1: Validierung
Rufe zuerst validate_offer auf und verschaffe dir einen Überblick.
prices_valid allein ist kein ausreichendes Abbruchkriterium. Prüfe immer auch die einzelnen Problemlisten, insbesondere duplicates und gp_errors.
Die Liste der Unsicherheiten ist keine eigenständige Arbeitsliste. Untersuche diese Positionen nur, wenn sie gleichzeitig in einer relevanten Problemliste auftauchen.
Phase 2: Probleme untersuchen
Arbeite die Probleme in dieser Reihenfolge ab:
1.	duplicates
2.	gp_errors
3.	orphans
4.	incomplete
Untersuche zusammenhängende Probleme gemeinsam. Verwende get_position möglichst für mehrere Positionen in einem Aufruf.
A. DUPLICATES
Eine gültige Positionsnummer kommt mehrfach im Extraktionsergebnis vor.
•	Vergleiche alle Einträge mit get_position. Sind mehrere Einträge visuell eindeutig derselben Soll-LV-Position zuzuordnen und kann keiner davon umbenannt werden, bleibt das Duplikat bestehen. Führe keine weiteren Rescans durch, nur um ein technisch nicht behebbares Duplikat aufzulösen. Dokumentiere es als ungeklärt.
•	Untersuche die ursprünglichen Bildinformationen mit rescan_page.
•	Stelle fest, ob ein Eintrag tatsächlich zu einer anderen Position gehört oder dieselbe Position mehrfach erkannt wurde.
•	Unterschiedliche Preise bedeuten nicht automatisch, dass einer falsch ist.
•	Korrigiere eine Positionsnummer nur, wenn der Scan die richtige Nummer und das passende Stichwort eindeutig bestätigt.
•	Lässt sich ein Duplikat nicht eindeutig auflösen, ändere nichts und melde es.
Mit den vorhandenen Tools können redundante Einträge nicht gelöscht werden.
B. GP_ERRORS
Ein Gesamtpreis stimmt rechnerisch nicht mit Einheitspreis mal Soll-LV-Menge überein.
•	Überprüfe Menge, EP und GP mit get_position.
•	Untersuche beide Preise unabhängig voneinander mit rescan_page.
•	Stimmen die extrahierten Werte mit dem Scan überein, ändere nichts. Der Bieter hat sich möglicherweise verrechnet.
•	Weicht ein Preis von der visuellen Erkennung ab, korrigiere ausschließlich diesen Wert.
•	Sind beide Werte falsch erkannt, korrigiere beide anhand des Scans.
•	Ist ein Wert nicht eindeutig lesbar, lasse ihn unverändert.
Wichtig: Du darfst niemals einen Gesamtpreis anhand von Einheitspreis mal Menge korrigieren. Auch ein mathematisch falscher GP bleibt unverändert, wenn er tatsächlich so im Angebot steht. Für die Weiterverarbeitung ist der Einheitspreis maßgeblich.
C. ORPHANS
Orphans sind Preise ohne Positionsnummer. Sie können beispielsweise durch Seitenumbrüche entstehen.
•	Vergleiche sie mit fehlenden oder unvollständigen Soll-LV-Positionen.
•	Suche nach einer plausiblen Zielposition und überprüfe die Zuordnung mit rescan_page.
•	Sind Positionsnummer und Preise eindeutig zugeordnet, aktualisiere den bereits vorhandenen Eintrag dieser Soll-LV-Position.
•	Ordne Preise niemals allein anhand ihrer Höhe oder Reihenfolge zu.
•	Ist keine eindeutige Zuordnung möglich, lasse den Orphan unverändert und melde ihn.
D. PRICES.INCOMPLETE
Hier fehlen Preisangaben bei einer Soll-LV-Position.
Unterscheide:
•	Preiszeile leer oder durchgestrichen: Der fehlende Wert ist korrekt. Nichts ändern und als nicht bepreist melden.
•	Preis beim Auslesen übersehen: Mit rescan_page überprüfen und eindeutig erkannte Werte eintragen.
•	Seitenumbruch: Beide Seiten untersuchen. Preise von der Folgeseite übernehmen, sofern sie eindeutig zur Position gehören.
•	Soll-LV-Menge fehlt: Keine Berechnung versuchen und nicht verändern.
•	Angebotspreis fehlt: Keine wiederholten Rescans durchführen. Im Bericht erwähnen.
Existiert für eine fehlende Position kein Eintrag im Extraktionsergebnis, kann update_position sie nicht ergänzen. Versuche keine unmögliche Änderung und melde sie als ungeklärt.
UNKNOWN UND MISSING
Nutze unbekannte Nummern als zusätzliche Hinweise bei der Untersuchung von Duplikaten, Orphans und fehlenden Positionen.
Vergleiche sie insbesondere mit fehlenden Soll-LV-Positionen. Eine ähnliche Nummer allein rechtfertigt keine Korrektur. Erst wenn rescan_page die tatsächliche Positionsnummer und das passende Stichwort bestätigt, darfst du einen bestehenden Eintrag umbenennen.
Unbekannte Nummern ohne plausible Verbindung zu einer Soll-LV-Position werden ignoriert und nicht einzeln berichtet.
Fehlende Positionen, die durch keine Korrektur wiederhergestellt werden können, bleiben ungeklärt und werden im Bericht aufgeführt.
4. RESCAN_PAGE
Übergib nur die Positionsnummer. Die Frage an das Vision-Modell wird automatisch gebaut.
Maximal zwei Rescans pro Position. Wiederhole keinen Rescan mit identischen Argumenten.
5. UPDATE_POSITION
Ändere Einträge ausschließlich auf Grundlage eindeutiger visueller Erkenntnisse aus rescan_page.
•	Verwende nur tatsächlich im Scan erkannte Werte.
•	Bei Duplikaten muss der richtige duplikat_index ausgewählt werden (beginnend bei 1, nicht mit dem globalen Index verwechseln).
•	Positionsnummern dürfen nur auf eine im Soll-LV existierende und im Scan bestätigte Nummer geändert werden.
•	Preise als normale numerische Werte übergeben (1234.5 statt 1.234,50).
•	Keine Bieterkommentare ändern.
•	Ist ein Problem eindeutig geklärt, setze Unsicherheit auf false.
•	Fehlgeschlagene Änderungen nicht mit identischen Argumenten wiederholen. Prüfe immer das Ergebnis des Tool-Aufrufs.
6. VALIDIERUNG UND ABSCHLUSS
•	Untersuche alle sinnvoll lösbaren Probleme und führe sichere Korrekturen durch.
•	Rufe validate_offer nach einem zusammenhängenden Bearbeitungsdurchlauf erneut auf, nicht nach jeder einzelnen Änderung.
•	Prüfe, ob die Änderungen erfolgreich waren und ob weitere Probleme bestehen.
•	Nach der letzten Änderung ist eine abschließende Validierung verpflichtend.
Eine Abweichung zwischen calculated_total und AI_scanned_total rechtfertigt keine weiteren Rescans oder Preisänderungen. Die Endsumme darf niemals verwendet werden, um einzelne Positionspreise abzuleiten oder zu korrigieren.
Arbeite mit einem Tool-Aufruf pro Nachricht und warte jeweils auf das Ergebnis. Kündige Tool-Aufrufe nicht in Textform an.
Werden durch die erneute Validierung weitere sinnvoll lösbare Probleme sichtbar, bearbeite diese in einem weiteren Durchlauf. Wiederhole die Prüfung, bis keine weiteren sicheren Korrekturen möglich sind. Vermeide wiederholte Durchläufe ohne neue Erkenntnisse. Beende die Prüfung, sobald alle sinnvoll lösbaren Probleme bearbeitet und alle verbleibenden Auffälligkeiten dokumentiert sind.

7. ABSCHLIESSENDER BERICHT
Schreibe einen deutschen Bericht mit maximal zehn Zeilen. Leere Abschnitte weglassen. Fasse mehrere Positionen desselben Problemtyps in einer Zeile zusammen. Die Grenze von zehn Zeilen hat Vorrang vor einer Einzelauflistung jeder Position.
1. Ungeklärt: Positionsnummer und kurzer Grund, weshalb keine sichere Korrektur möglich war.
2. Ansehen: Auffälligkeiten, die der Bearbeiter prüfen sollte, beispielsweise tatsächliche Rechenfehler oder nicht eingetragene Preise.
3. Wichtig: Abweichung zwischen Angebotspreis und berechneter Summe mit beiden Werten und Differenz, fehlende Positionen und nicht bepreiste Positionen.
Zähle vorgenommene Änderungen nicht auf. Diese werden separat protokolliert. Erwähne keine Bieterkommentare oder Unsicherheitsmarkierungen. Stelle Vermutungen nicht als Tatsachen dar.
Wenn keine relevanten Auffälligkeiten bestehen, schreibe:
"Die Prüfung ergab keine offenen Auffälligkeiten."

"""
RESCAN_MODEL = "qwen3-vl:4b-instruct"   # ein anderes Modell als in LLM_call eintragen
LLM_CALL_MODEL = "qwen3-vl:4b-instruct"
AGENT_MODEL =  "qwen3-vl:4b-instruct"

RESCAN_PROMPT_AUSSCHNITT = """Die Bilder zeigen Ausschnitte eines Leistungsverzeichnisses, in das der Bieter Preise von Hand eingetragen hat. Das erste Bild beginnt mit der Position {position} (Stichwort: {stichwort}). Falls ein zweites Bild dabei ist, zeigt es den Anfang der Folgeseite.
Nenne zuerst die gedruckte Positionsnummer und das Stichwort exakt wie im Bild. Steht die Position nicht im Bild, antworte 'nicht gefunden' und rate nicht.
Die Preiszeile der Position steht nach ihrem Positionstext und vor der nächsten Positionsnummer. Preise, die nach einer anderen Positionsnummer stehen, gehören nicht dazu.
Nenne den handschriftlichen Einheitspreis (Zeile 'Einheitspreis') und den handschriftlichen Gesamtpreis (rechts auf derselben Zeile nach 'EUR') unabhängig voneinander. Rechne nichts nach.
Gib ausdrücklich an, wenn ein Preis leer, durchgestrichen, korrigiert oder nicht lesbar ist. Erfinde keine Informationen."""

RESCAN_PROMPT_SEITE = """Die Bilder zeigen zwei aufeinanderfolgende Seiten eines Leistungsverzeichnisses, in das der Bieter Preise von Hand eingetragen hat.
Suche die Position {position} (Stichwort: {stichwort}).
Nenne zuerst die gedruckte Positionsnummer und das Stichwort exakt wie im Bild. Steht die Position nicht auf den Seiten, antworte 'nicht gefunden' und rate nicht.
Die Preiszeile einer Position steht UNTER ihrem Positionstext, nie darüber. Preise, die über der Positionsnummer stehen, gehören zur vorherigen Position.
Nenne den handschriftlichen Einheitspreis (Zeile 'Einheitspreis') und den handschriftlichen Gesamtpreis (rechts auf derselben Zeile nach 'EUR') unabhängig voneinander. Rechne nichts nach.
Gib ausdrücklich an, wenn ein Preis leer, durchgestrichen, korrigiert oder nicht lesbar ist. Erfinde keine Informationen."""



BASE_DIR = Path(__file__).resolve().parent
TESSDATA = str(BASE_DIR / "tessdata")
keywords = ["Einheitspreis", "Lohn"]
ERLAUBTE_FELDER = {"positionsnummer", "einheitspreis", "gesamtpreis", "Unsicherheit"}
NUMMER_MUSTER = re.compile(r"\d{2}(?:\.\d{2})+(?:\.[A-Za-z0-9])?")
IMMER_BEHALTEN = {"1"}

rescan_count = {}

def t(name):                       
    return f"{{{URI}}}{name}"

def parse_xml(file_path):
    Soll_LV = {}
    tree = ET.parse(file_path)
    root = tree.getroot()

    global URI, NS
    URI = root.tag.split("}")[0][1:]   
    NS = {"on": URI}                   
    ET.register_namespace("", URI)


    og, lg, ulg, gnr, ftnr = None, None, None, None, None

    for element in root.iter():
        if element.tag == t("og"): 
            og = element.attrib['nr']
            lg, ulg, gnr, ftnr = None, None, None, None
        if element.tag == t('lg'): 
            lg = element.attrib['nr']
            ulg, gnr, ftnr = None, None, None
        if element.tag == t('ulg'): 
            ulg = element.attrib['nr']
            gnr, ftnr = None, None
        if element.tag == t('grundtextnr'): 
            gnr = element.attrib['nr']
            ftnr = None
        if element.tag == t('folgeposition'): 
            ftnr = element.attrib['ftnr']
        elif element.tag == t('ungeteilteposition'):
            ftnr = None
        if element.tag == t('pos-eigenschaften'):
            lvmenge = element.findtext(t('lvmenge'))
            if lvmenge is not None:
                description = element.findtext(t('stichwort'))
                einheit = element.findtext(t('einheit'))
                positionsnummer = [og, lg, ulg, gnr, ftnr]
                positionsnummer = [num for num in positionsnummer if num is not None]
                positionsnummer = '.'.join(positionsnummer)
                position = Position(positionsnummer, description, lvmenge, einheit, element)
                if positionsnummer not in Soll_LV: 
                    Soll_LV[positionsnummer] = position 
                else:
                    print(f'Warning: Duplicate position number {positionsnummer} found.')
                    raise ValueError(f'Duplicate position number found: {positionsnummer}. A. {Soll_LV[positionsnummer].description} B. {description}')
    return Soll_LV, tree

def excel_LV(file_path):
  Soll_LV = {}
  #indexes = {}
  workbook = openpyxl.load_workbook(file_path)
  sheet = workbook["Sheet1"]

  for row in range(1, sheet.max_row + 1):
      if sheet.cell(row=row, column=4).value == 'P':
          pos_number = sheet.cell(row=row, column=1).value
          description = sheet.cell(row=row, column=3).value
          menge = sheet.cell(row=row, column=12).value
          einheit = sheet.cell(row=row, column=13).value
          Soll_LV[pos_number] = Position(pos_number, description, menge, einheit)
  return Soll_LV

def build_soll_lv_context(Soll_LV):
    referenz_positionen = []

    for nummer, position in Soll_LV.items():
        referenz_positionen.append({
            "positionsnummer": nummer,
            "beschreibung": position.description,
            "lvmenge": position.lvmenge,
            "einheit": position.einheit
        })

    return json.dumps(
        referenz_positionen,
        ensure_ascii=False,
        indent=2
    )

def seiten_text(page, Soll_LV):
    page.remove_rotation()
    text = page.get_text()
    position_cords = []
    position_list = []

    current_x = []

    if len(text.strip()) < 50:
      textpage = page.get_textpage_ocr(language="deu", dpi=300, full=True, tessdata=TESSDATA)
      text = textpage.extractText()
    else:
        textpage = page.get_textpage()

    words = page.get_text("words", textpage=textpage)
    for w in words:
        wort = w[4].strip(".,;:")          # OCR hängt manchmal Satzzeichen an
        if wort in Soll_LV:                # exakter Treffer im Dictionary
            bbox = pymupdf.Rect(w[0], w[1], w[2], w[3])
            position_cords.append(bbox)
            position_list.append(wort)

      #schauen ob alle in einer linie sind wenn nicht außreiser entfernen
    
    paare = list(zip(position_cords, position_list)) #liste kombinieren um sie zu sortieren nach y
    paare.sort(key=lambda p: p[0].y0)

    #Listen wieder trennen
    position_cords = [p[0] for p in paare]
    position_list = [p[1] for p in paare]

    if len(position_cords) > 2:
        for pos in position_cords:
            avg_x = (pos.x0 + pos.x1) / 2
            current_x.append(avg_x)

        median_x = numpy.median(current_x)
        tolerance = 30
        filtered_cords = []
        filtered_positions = []

        for i, x in enumerate(current_x):
            if abs(x - median_x) <= tolerance:
                filtered_cords.append(position_cords[i])
                filtered_positions.append(position_list[i])
            else:
                print(
                    f"Ausreißer entfernt: {position_list[i]} "
                    f"bei X={x}"
                )
                
        position_cords = filtered_cords
        position_list = filtered_positions

      # den oberen Crop bis zu ersten Position berücksichtigen
    position_list.insert(0, "eventuell keine oder eine abgeschnittene Position")
    return text, position_cords, position_list

def position_auf_seite(page, Soll_LV):
    text, position_cords, position_list = seiten_text(page, Soll_LV)
    score = 0
    for word in keywords:
        if word in text:
            score += 3
    #         print(word)
    # print(f"score nach wörtern = {score}")
    for key in Soll_LV:
        if key in text:
            score += 5
    #         print(key)
    # print(f"score nach keys = {score} \n")
    return (5 <= score), text, position_cords, position_list

def pdf_verarbeitung(pdf_path, Soll_LV):
    with pymupdf.open(pdf_path) as doc:
        docs = []
        found = False
        text_data = []
        cropped_images = {}
        positions_in_image = {}

        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            pos_on_page, text, position_cords, position_list = position_auf_seite(page, Soll_LV)

            if found or pos_on_page:
                pix = page.get_pixmap(dpi=300)

                script_dir = Path(__file__).resolve().parent
                temp_dir = script_dir / 'temp'
                temp_dir.mkdir(parents=True, exist_ok=True)
                output_filename = temp_dir / f'seite_{page_num}.png'
                pix.save(str(output_filename))

                docs.append(str(output_filename))
                text_data.append(text)

                index = len(docs) -1
                positions_in_image[index] = position_list
                cropped_images[index] = (speichere_seitenbereiche(
                  page,
                  position_cords,
                  index,
                  ausgabe_ordner=str(temp_dir)
                ))
                found = True

    return docs, text_data, cropped_images, positions_in_image

def speichere_seitenbereiche(
    page,
    positions,
    index,
    ausgabe_ordner="ausschnitte",
    dpi=200
):
    """
    Schneidet eine PDF-Seite anhand der gefundenen Wortpositionen
    in mehrere Bereiche und speichert diese als PNG-Bilder.

    :param page: PyMuPDF Page-Objekt
    :param positions: Liste von pymupdf.Rect mit Wortpositionen
    :param ausgabe_ordner: Ausgabeordner für die Bilder
    :param dpi: Auflösung der gespeicherten Bilder
    :return: Liste der gespeicherten Bildpfade
    """

    os.makedirs(ausgabe_ordner, exist_ok=True)

    # Positionen anhand der Y-Koordinate sortieren
    positions = sorted(positions, key=lambda rect: rect.y0)

    seite = page.rect

    # Schnittgrenzen festlegen
    grenzen = [seite.y0]

    for pos in positions:
        grenzen.append(pos.y0)

    grenzen.append(seite.y1)

    grenzen = sorted(set(grenzen))

    gespeicherte_bilder = []

    # Bereiche ausschneiden und als Bild speichern
    for i in range(len(grenzen) - 1):

        y_oben = grenzen[i] - 10
        y_unten = grenzen[i + 1]

        if y_unten <= y_oben:
            continue

        clip = pymupdf.Rect(
            seite.x0,
            y_oben,
            seite.x1,
            y_unten
        )

        # Bereich als Pixmap rendern
        pix = page.get_pixmap(
            matrix=pymupdf.Matrix(dpi / 72, dpi / 72),
            clip=clip,
            alpha=False
        )

        dateiname = os.path.join(
            ausgabe_ordner,
            f"index_{index}_ausschnitt_{i + 1}.png"
        )
        pix.save(dateiname)

        gespeicherte_bilder.append(dateiname)

    return gespeicherte_bilder

def Mögliche_Positionen_pydantic(Soll_LV, positionen):
    return {
        key: Soll_LV[key]
        for key in positionen
        if key in Soll_LV
    }

#def multistep_model_call():

def LLM_call(Soll_LV, cropped_images, positions_list, docs):
    if not docs:
        raise ValueError("Es wurden keine LV-Seiten im PDF erkannt.")
    responses = []
    response_sources = []
    for i in range(len(cropped_images)):
        selected_positions = build_soll_lv_context(Mögliche_Positionen_pydantic(Soll_LV, positions_list[i]))
        
        for index, image in enumerate(cropped_images[i]):
            
            most_likely = positions_list[i][index]
            response = ollama.chat(
                model=LLM_CALL_MODEL,
                format=positionList.model_json_schema(),
                messages=[
                    {
                        "role": "user",
                        "content": PROMPT + "\nWahrscheinlichste Position (nur Richtwert, kann falsch sein):" + most_likely + "\nSoll-LV-Positionen dieser Seite (Hilfe zum Abgleichen): \n" + selected_positions,
                        "images": [image],
                    }
                ],
                options={"temperature": 0, "num_ctx": 8192},
            )
            print(response['message']['content'])
            try:
                responses.append(json.loads(response["message"]["content"]))
                response_sources.append({
                    "crop_path": image,
                    "page_path": docs[i],
                    "page_index": i,
                    "crop_index": index,
                    "is_last_crop": index == len(cropped_images[i]) - 1,
                })
            except Exception as ex:
                print("json loads error im LLM call", ex)

    return responses, response_sources

def finde_ausschnitte(position, duplikat_index=1):
    """Bildpfade für den Rescan, oder [] wenn es zu der Nummer keinen Eintrag gibt."""
    global position_sources, cropped_images
    treffer = [i for i, pos in enumerate(pydanticobj.positions) if pos.positionsnummer == position]
    if not 1 <= duplikat_index <= len(treffer):
        return []
    quellen = position_sources[treffer[duplikat_index - 1]]
    bilder = [q["crop_path"] for q in quellen]

    # Nicht zusammengefügt und letzter Ausschnitt der Seite: Anfang der Folgeseite dazu
    letzte = quellen[-1]
    if len(quellen) == 1 and letzte["is_last_crop"] and cropped_images is not None:
        folgeseite = cropped_images.get(letzte["page_index"] + 1)
        if folgeseite:
            bilder.append(folgeseite[0])
    return bilder

def rescan_page(position: str, duplikat_index: int = 1) -> str | dict:
    """Lässt ein Vision-Modell den Bildausschnitt einer Position neu lesen und gibt seine Antwort als Text zurück. Die Frage an das Modell wird automatisch gebaut. Höchstens 2 Neuscans pro Position.
        Args:
          position: Positionsnummer, die neu gelesen werden soll
          duplikat_index: bei Duplikaten der wievielte Eintrag (beginnend bei 1)"""
    global docs, text
    count = rescan_count.get(position, 0)
    if count >= 2:
        return {"success": False,
                "error": "Maximale Anzahl an Rescans erreicht. Lass diese Position in Ruhe und bearbeite andere Problemstellen."}
    rescan_count[position] = count + 1

    stichwort = (Soll_LV[position].description if position in Soll_LV else None) or "unbekannt"
    images = finde_ausschnitte(position, duplikat_index)

    if images:
        prompt = RESCAN_PROMPT_AUSSCHNITT.format(position=position, stichwort=stichwort)
    else:
        # Rückfall: ganze Seiten wie bisher
        pattern = re.compile(rf"(?<![\w.]){re.escape(position)}(?![\w.])")
        target_page = next((p for p in range(len(text)) if pattern.search(text[p])), -1)
        if target_page == -1:
            return f"Page with position={position} was not found"
        if docs is not None: images = docs[target_page:target_page + 2]
        else: raise ValueError("docs ist nicht definiert,")
        prompt = RESCAN_PROMPT_SEITE.format(position=position, stichwort=stichwort)

    response = ollama.chat(
        model=RESCAN_MODEL,
        messages=[{"role": "user", "content": prompt, "images": images}],
        options={"temperature": 0, "num_ctx": 16384},
    )
    return response["message"]["content"]

def update_position(position_id: str, field: str, new_value: str | float | bool | None,
                    reason: str, duplikat_index: int = 1) -> dict: # pydanticobj, history müssen globale Variablen sein dann hat das LLM nicht so viel zu bestimmen
    """Ändert ein Feld einer ausgelesenen Position. Gibt den alten und neuen Wert zurück, oder einen Fehler, wenn Position oder Feld nicht gefunden wurden.
      Args:
        position_id: aktuelle Positionsnummer des Eintrags
        field: "positionsnummer", "einheitspreis", "gesamtpreis" oder "Unsicherheit"
        new_value: neuer Wert (Text, Zahl oder true/false)
        reason: Begründung, was rescan_page gezeigt hat
        duplikat_index: Für duplikate um gezielt z.b. das 2te zu bekommen"""
    global pydanticobj, history, Soll_LV
    dic = pydanticobj.model_dump()
    if field not in ERLAUBTE_FELDER:
        return {
            "action": "update_position",
            "success": False,
            "error": f"Feld '{field}' ist nicht erlaubt. Erlaubt: {sorted(ERLAUBTE_FELDER)}"
        }
    
    report = {"action": "update_position", "success": False, "error": ""}

    matches = [
        i for i, pos in enumerate(dic["positions"])
        if pos["positionsnummer"] == position_id
    ]

    if not 1 <= duplikat_index <= len(matches):
        return {
            "action": "update_position",
            "success": False,
            "error": "Position oder Duplikat-Index nicht gefunden."
        }

    index = matches[duplikat_index - 1]

    if field in dic["positions"][index]:
        old_value = dic["positions"][index][field]
        try:
            if field == "positionsnummer" and (not isinstance(new_value, str) or new_value not in Soll_LV):
                return {"action": "update_position", "success": False, "error": "Ungültige Positionsnummer."}

            if field == "Unsicherheit" and not isinstance(new_value, bool):
                return {"action": "update_position", "success": False, "error": "Unsicherheit muss bool sein."}

            if field in {"einheitspreis", "gesamtpreis"} and new_value is not None and (isinstance(new_value, bool) or not isinstance(new_value, (int, float)) or not numpy.isfinite(new_value)):
                return {"action": "update_position", "success": False, "error": "Preis muss eine endliche Zahl sein."}

            dic["positions"][index][field] = new_value

            temp = FinalResult(**dic)  #erster Validierung probieren und dann globale Variable überschreiben
            pydanticobj = temp

            report = {
                "action": "update_position",
                "success": True,
                "position_number": position_id,
                "field": field,
                "old_value": old_value,
                "new_value": new_value,
                "reason": reason
            }

            history.append({
                "action": "update_position",
                "position_id": position_id,
                "field": field,
                "old_value": old_value,
                "new_value": new_value,
                "reason": reason
            })
        except Exception as e:
            report= {
                "action": "update_position",
                "success": False,
                "error": str(e),
            }

    return report

def get_position(positions: list[str]) -> list[dict]:
    """Zeigt zu jeder Positionsnummer die Soll-LV-Daten (Beschreibung, Menge, Einheit) und alle bisher ausgelesenen Einträge mit index, Preisen und Unsicherheit. Funktioniert auch für Nummern, die nicht im Soll-LV stehen (im_soll_lv ist dann false).
    Args:
      positions: Liste von Positionsnummern, exakt so geschrieben wie in validate_offer angezeigt"""

    global pydanticobj
    dic = pydanticobj.model_dump()
    pos_info = []

    for pos in positions:

        # Alle extrahierten Einträge dieser Nummer suchen
        extracted_entries = []

        for i, json_pos in enumerate(dic["positions"]):
            if json_pos["positionsnummer"] == pos:
                extracted_entries.append({
                    "index": i,
                    "positionsnummer": json_pos["positionsnummer"],
                    "einheitspreis": json_pos["einheitspreis"],
                    "gesamtpreis": json_pos["gesamtpreis"],
                    "Unsicherheit": json_pos["Unsicherheit"],
                    "bieterkommentar": json_pos["bieterkommentar"]
                })

        # Soll-LV-Informationen abrufen, falls vorhanden
        if pos in Soll_LV:
            soll_lv = {
                "beschreibung": Soll_LV[pos].description,
                "menge": Soll_LV[pos].lvmenge,
                "einheit": Soll_LV[pos].einheit
            }
        else:
            soll_lv = None

        # Ergebnis zusammenstellen
        pos_info.append({
            "angefragte_positionsnummer": pos,
            "im_soll_lv": pos in Soll_LV,
            "soll_lv": soll_lv,
            "extracted_entries": extracted_entries
        })

    return pos_info

def TeilA_oder_TeilB(pos):
    EP = pos.einheitspreis
    GP = pos.gesamtpreis
    pos_num = pos.positionsnummer
    if pos_num is not None and EP is None and GP is None:
        return "A", pos
    elif (pos_num is None or pos_num in ["0", "", "1"]) and (EP is not None or GP is not None):
        return "B", pos
    return "C", pos

def create_finaljson(responses, response_sources):
    if len(responses) != len(response_sources):
        raise ValueError("Responses und Sources haben unterschiedliche Längen.")
    
    all_positions = []
    position_sources = []

    angebots_preis = None
    kommentare = []

    last_half = "C"
    last_pos = None

    for response, source in zip(responses, response_sources):
        try:
            parsed = positionList.model_validate(response)
        except ValidationError as e:
            kommentare.append(f"Antwort nicht validierbar ({e.error_count()} Fehler), übersprungen.")
            last_half = "C"
            continue

        if len(parsed.positions) != 0:
            this_half, this_pos = TeilA_oder_TeilB(parsed.positions[0])

            if this_half == "B" and last_half == "A":
                if last_pos is not None:
                    last_pos.einheitspreis = this_pos.einheitspreis
                    last_pos.gesamtpreis = this_pos.gesamtpreis

                    if last_pos.bieterkommentar is None and this_pos.bieterkommentar is not None:
                        last_pos.bieterkommentar = this_pos.bieterkommentar

                    last_pos.Unsicherheit = False

                    position_sources[-1].append(source)

                    kommentare.append(f"Positionen nach Seitenumbruch vereint: {last_pos}")
                    parsed.positions = parsed.positions[1:]

            if len(parsed.positions) != 0:
                last_half, last_pos = TeilA_oder_TeilB(parsed.positions[-1]) # falls auf der Seite genau ein Teil B ist
            else: last_half = "C"
        else:
            last_half = "C"

        all_positions.extend(parsed.positions)
        position_sources.extend([[source.copy()] for _ in parsed.positions])

        if parsed.angebots_preis is not None:
            angebots_preis = parsed.angebots_preis

        if parsed.AIkommentar:
            kommentare.append(parsed.AIkommentar)

    final_result = FinalResult(
        positions=all_positions,
        angebots_preis=angebots_preis,
        Bericht=" ".join(kommentare)
    )

    return final_result, position_sources

def validate_GP_AP(json_outputs, Soll_LV, duplicates, missing):
    GP_errors = []
    incomplete_pos = []
    angebots_preis = Decimal("0.0")
    AI_AP = None

    for json_pos in json_outputs["positions"]:
        pos_number = json_pos["positionsnummer"]

        if pos_number in Soll_LV:

            lvmenge = Soll_LV[pos_number].lvmenge
            Einheitspreis = json_pos["einheitspreis"]
            GP_AI = json_pos["gesamtpreis"]

            if lvmenge is None or Einheitspreis is None:
                incomplete_pos.append(pos_number)
            else:
                GP_python = Decimal(str(lvmenge)) * Decimal(str(json_pos["einheitspreis"]))
                if GP_AI is not None:
                    GP_AI = Decimal(str(GP_AI))
                    
                    if abs(GP_python - GP_AI) > Decimal("0.1"):
                        GP_errors.append(pos_number)

                if pos_number not in duplicates:
                    angebots_preis += GP_python

    try:
        AI_AP = Decimal(str(json_outputs["angebots_preis"]))
    except Exception:
        print("hat nicht funktioniert mit An Preis")

    if json_outputs["angebots_preis"] is None:
        incomplete_pos.append("angebots_preis")
    else:
        if (not duplicates and not missing and abs(angebots_preis - Decimal(str(json_outputs["angebots_preis"]))) < Decimal("1.00") and len(incomplete_pos) == 0):
            return True, GP_errors, incomplete_pos, angebots_preis, AI_AP
               
    return False, GP_errors, incomplete_pos, angebots_preis, AI_AP

def filtere_unknown(unknown, soll_nummern):
    """Gibt (echte Verdachtsfälle, Anzahl aussortierter Einträge) zurück."""
    soll_nummern = list(soll_nummern)
    unknown = list(dict.fromkeys(unknown))  # duplicates raus filtern

    behalten = []
    for nummer in unknown:
        if nummer in IMMER_BEHALTEN:
            behalten.append(nummer)
            continue

        if not NUMMER_MUSTER.fullmatch(nummer):
            continue
        
        if any(s.startswith(nummer + ".") for s in soll_nummern): # schauen ob es der anfang einer Nummer ist.
            continue
        behalten.append(nummer)

    return behalten, len(unknown) - len(behalten)

def validate_positions(json_out, Soll_LV):
    existing_pos = []
    missing_pos = []
    unknown_pos = []
    unsicherheiten = []
    duplicates = {}
    counter = {}
    extracted_positions = [
        pos["positionsnummer"]
        for pos in json_out["positions"]
        if pos["positionsnummer"] is not None
    ]
    for pos in Soll_LV:
        if pos in extracted_positions: 
            existing_pos.append(pos)
        else:
            missing_pos.append(pos)

    for pos in extracted_positions:
        if pos not in counter:
            counter[pos] = 0
        if pos in Soll_LV:
                counter[pos] += 1
        else:
            unknown_pos.append(pos)
    for x in counter:
        if counter[x] > 1:
            duplicates[x] = counter[x]

    orphans = [
        {
            "index": i,
            "einheitspreis": pos["einheitspreis"],
            "gesamtpreis": pos["gesamtpreis"],
            "Unsicherheit": pos["Unsicherheit"]
        }
        for i, pos in enumerate(json_out["positions"])
        if pos["positionsnummer"] is None
        and pos["bieterkommentar"] is None
        and (
            pos["einheitspreis"] is not None
            or pos["gesamtpreis"] is not None
        )
    ]

    unsicherheiten = [
        pos["positionsnummer"]
        for pos in json_out["positions"]
        if pos["Unsicherheit"] is True
    ]

    return existing_pos, missing_pos, unknown_pos, duplicates, orphans, unsicherheiten

def validate_offer():
    """Prüft den aktuellen Stand der ausgelesenen Positionen gegen das Soll-LV und meldet alle Probleme. Keine Parameter. Rufe es zuerst auf und nach jeder Änderung erneut.
    Rückgabe:
    - positions.existing: korrekt gefundene Nummern
    - positions.missing: Nummern aus dem Soll-LV, die im Angebot nicht gefunden wurden
    - positions.unknown: gelesene Nummern, die nicht im Soll-LV stehen.
    - positions.duplicates: Nummern, die mehrfach vorkommen, mit Anzahl
    - positions.orphans: Einträge ohne Positionsnummer, aber mit Preisen, jeweils mit index
    - positions.Unsichere Positionen: Positionen, bei denen sich das Vision-Modell unsicher war
    - prices.prices_valid: true, wenn berechnete Summe und gelesene Endsumme zusammenpassen und nichts unvollständig ist
    - prices.gp_errors: Positionen, deren Gesamtpreis nicht zu Einheitspreis mal Menge passt
    - prices.incomplete: Positionen, bei denen Preise fehlen; enthält "angebots_preis", wenn die Endsumme nicht gelesen wurde
    - prices.calculated_total: aus Einheitspreis mal Menge berechnete Summe
    - prices.AI_scanned_total: vom Vision-Modell gelesene Endsumme
    - prices.Differenz: Betrag der Abweichung zwischen berechneter und gelesener Summe
    """
    global pydanticobj
    json_outputs = pydanticobj.model_dump()

    existing, missing, unknown, duplicates_dic, orphans, unsicherheiten = validate_positions(json_outputs, Soll_LV)
    unknown, _ = filtere_unknown(unknown, Soll_LV.keys())
    richtiger_AP, GP_errors, incomplete_pos, calculated_total, AI_total = validate_GP_AP(json_outputs, Soll_LV, duplicates_dic, missing)

    Differenz = None
    if calculated_total is not None and AI_total is not None:
        Differenz = abs(calculated_total - AI_total)
    return {
        "positions": {
            "duplicates": duplicates_dic,
            "existing": len(existing),
            "missing": missing,
            "unknown": unknown,
            "orphans": orphans,
            "Unsichere Positionen": unsicherheiten
        },
        "prices": {
            "prices_valid": richtiger_AP,
            "gp_errors": GP_errors,
            "incomplete": incomplete_pos,
            "calculated_total": calculated_total,
            "AI_scanned_total": AI_total,
            "Differenz": Differenz
        }
    }

available_functions = {
    "rescan_page": rescan_page,
    "update_position": update_position,
    "validate_offer": validate_offer,
    "get_position": get_position
}

def run_agent(system_prompt):

    messages = [
        {
            "role": "system",
            "content": system_prompt
        },
        {
            "role": "user",
            "content": "Beginne bitte mit der Bearbeitung. Du darfst nur 1 Tool pro aufruf verwenden. Starte mit validate_offer um einen Überblick zu bekommen."
        }
    ]

    max_iterations = 50

    for iteration in range(max_iterations):

        response = ollama.chat(
            model=AGENT_MODEL,
            messages=messages,
            tools=[rescan_page, validate_offer, update_position, get_position],
            options={"temperature": 0, "num_ctx": 16384},
        )

        # Antwort des Modells zum Gespräch hinzufügen
        assistant_message = response.message
        tool_calls = assistant_message.tool_calls or []

        # 1. Tool-Aufrufe begrenzen
        MAX_TOOLS_PER_TURN = 1

        if len(tool_calls) > MAX_TOOLS_PER_TURN:
            print(f"Modell wollte {len(tool_calls)} Tools rufen. "
                  f"Limitiere auf {MAX_TOOLS_PER_TURN}.")

            tool_calls = tool_calls[:MAX_TOOLS_PER_TURN]
            assistant_message.tool_calls = tool_calls

        # Assistant-Nachricht erst nach der Begrenzung speichern
        messages.append(assistant_message)

        if not tool_calls:
            return assistant_message.content

        # Tool-Aufrufe ausführen
        for tool_call in tool_calls:

            function_name = tool_call.function.name
            arguments = tool_call.function.arguments

            print(f"Agent ruft auf: {function_name}")
            print(f"Argumente: {arguments}\n")

            function_to_call = available_functions.get(function_name)

            # Alte Validierungsergebnisse entfernen
            # if function_name == "validate_offer":

            #     messages = [
            #         msg for msg in messages
            #         if not (
            #             isinstance(msg, dict)
            #             and msg.get("role") == "tool"
            #             and msg.get("tool_name") == "validate_offer"
            #         )
            #     ]
            
            if function_to_call is None:
                result = {
                    "success": False,
                    "error": f"Unbekanntes Tool: {function_name}"
                }
            else:
                try:
                    result = function_to_call(**arguments)
                    print(result)
                except Exception as e:
                    result = {
                        "success": False,
                        "error": str(e)
                    }

            # Ergebnis an das Modell zurückgeben
            messages.append({
                "role": "tool",
                "tool_name": function_name,
                "content": json.dumps(
                    result,
                    ensure_ascii=False,
                    default=str
                )
            })
    #raise RuntimeError("Agent hat maximale Iterationen erreicht.")
    raise RuntimeError(
        "Supervisor-Agent hat das Iterationslimit erreicht. "
        "Prüfung wurde nicht vollständig abgeschlossen."
    )

def get_comments(result):
    bieterkommentare = []
    for i, pos in enumerate(result.positions):
        if pos.bieterkommentar:
            bieterkommentare.append({
                #"index": i,
                "positionsnummer": pos.positionsnummer,
                "kommentar": pos.bieterkommentar
            })
    return bieterkommentare


def create_LV(pydanticobj, Soll_LV, validation_json):
    json_outputs = pydanticobj.model_dump()

    LV = deepcopy(Soll_LV)
    for key in LV:
        LV[key].element = Soll_LV[key].element

    for positionLV in LV:
        position = LV[positionLV]

        for positionOutput in json_outputs["positions"]:
            pos_number = positionOutput["positionsnummer"]

            if pos_number == positionLV and pos_number not in validation_json["positions"]["duplicates"]: # schauen ob die position existiert oder zweimal da ist.

                position.einheitspreis = positionOutput["einheitspreis"]

                if position.einheitspreis is None:
                    continue

                if position.einheitspreis is not None and position.lvmenge is not None:
                    position.gesamtpreis = (
                        Decimal(str(position.einheitspreis))
                        * Decimal(str(position.lvmenge))
                    )
    return LV

def write_excel(final_LV, ausgang, file_name, input_excel):
    workbook = openpyxl.load_workbook(input_excel)
    sheet = workbook["Sheet1"]

    for position in final_LV:
        positionsnummer = final_LV[position].positionsnummer
        einheitspreis = final_LV[position].einheitspreis
        gesamtpreis = final_LV[position].gesamtpreis

        for row in range(1, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=1)

            if cell.value == positionsnummer:
                sheet.cell(row=row, column=16).value = einheitspreis
                sheet.cell(row=row, column=18).value = gesamtpreis

    datei_pfad = ausgang / f"{file_name}_excel.xlsx"
    workbook.save(datei_pfad)

def create_angebotsLV(final_LV, ausgang, file_name, tree):
    root = tree.getroot()
    for element in root.findall(t("ausschreibungs-lv")):
        element.tag = t("angebots-lv")

        summenormalausfuehrung = ET.SubElement(element, t("summenormalausfuehrung"))

        ET.SubElement(summenormalausfuehrung, t("summe-des-lv")).text = "0.00"
        ET.SubElement(summenormalausfuehrung, t("summe-nachlaesse")).text = "0.00"
        ET.SubElement(summenormalausfuehrung, t("gesamtpreis")).text = "0.00"

        ust = ET.SubElement(summenormalausfuehrung, t("ust"))
        ET.SubElement(ust, t("ustsatz")).text = "20.00"
        ET.SubElement(ust, t("ustbetrag")).text = "0.00"

        ET.SubElement(summenormalausfuehrung, t("angebotspreis")).text = "0.00"

    for element in root.iter():
        if element.tag == t("kenndaten"):
            auftraggeber = element.find(t("auftraggeber"))

            if auftraggeber is not None:
                AN_Abschnitt = deepcopy(auftraggeber)
                AN_Abschnitt.tag = t("bieter_auftragnehmer")

                element.append(AN_Abschnitt)

    for key in final_LV:
        element = final_LV[key].element
        nicht_angeboten = element.find(t("nichtangeboten"))

        if nicht_angeboten is not None and (final_LV[key].gesamtpreis is not None or final_LV[key].einheitspreis is not None):

            pzzv = element.find(t("pzzv"))
            ist_eventual = (
                pzzv is not None
                and pzzv.find(t("eventualposition")) is not None
            )

            if final_LV[key].gesamtpreis is None:
                gesamtpreis = "0.00"
            else:
                gesamtpreis = str(final_LV[key].gesamtpreis)

            if final_LV[key].einheitspreis is None:
                einheitspreis = "0.00"
            else:
                einheitspreis = str(final_LV[key].einheitspreis)

            index = list(element).index(nicht_angeboten)
            element.remove(nicht_angeboten)

            preis_elem = ET.Element(t("preis"))
            #ET.SubElement(preis_elem, t("preisanteil1")).text = "0.00"
            #ET.SubElement(preis_elem, t("preisanteil2")).text = einheitspreis
            ET.SubElement(preis_elem, t("gesamt")).text = einheitspreis

            pospreis_elem = ET.Element(t("pospreis"))
            pospreis_elem.text = gesamtpreis

            element.insert(index, preis_elem)
            element.insert(index + 1, pospreis_elem)

    ET.indent(root, space="  ")
    datei_pfad = ausgang / f"{file_name}_Angebots-ONLV.ONLV"
    tree.write(datei_pfad, encoding="utf-8", xml_declaration=True)

def delete_images(docs, cropped):
    if docs is not None:
        for file in docs:
            try:
                os.remove(file)
            except FileNotFoundError:
                print(f"Datei '{file}' wurde nicht gefunden.")
            except PermissionError:
                print("Keine Berechtigung zum Löschen der Datei.")
    if cropped is not None:
        for i in range(len(cropped)):
            for image in cropped[i]:
                try:
                    os.remove(image)
                except FileNotFoundError:
                    print(f"Datei '{image}' wurde nicht gefunden.")
                except PermissionError:
                    print("Keine Berechtigung zum Löschen der Datei.")

def complete_workflow(Input_LV, Dateiformat, Scans_dict, Ausgang, on_progress):
    global Soll_LV, pydanticobj, history, rescan_count, docs, text, position_sources, cropped_images
    history, failed = [], []
    rescan_count.clear()
    name, tree = None, None
    
    Scans = list(Scans_dict.values())
    
    on_progress(0.05, "Verarbeiten der LV Positionen")

    gesamt = len(Scans) * 4
    for i, pdf in enumerate(Scans):
        docs, cropped_images = None, None
        try:
            name = next((k for k, v in Scans_dict.items() if v == pdf), None)
            if name is not None: name = Path(name).stem

            if Dateiformat == "Excel":
                Soll_LV = excel_LV(Input_LV)
            elif Dateiformat == "ONLV":
                Soll_LV, tree = parse_xml(Input_LV)
            else:
                raise ValueError("Falsches Dateiformat")
        
            on_progress(0.05 + 0.95 * (i * 4 + 0) / gesamt, f"Anbieter {i + 1} von {len(Scans)}: PDF wird eingelesen")

            docs, text, cropped_images, position_list = pdf_verarbeitung(pdf, Soll_LV)

            on_progress(0.05 + 0.95 * (i * 4 + 1) / gesamt, f"Anbieter {i + 1} von {len(Scans)}: Positionen werden gelesen")

            responses, response_sources = LLM_call(Soll_LV, cropped_images, position_list, docs)

            pydanticobj, position_sources = create_finaljson(responses, response_sources)

            validation_data = validate_offer()

            on_progress(0.05 + 0.95 * (i * 4 + 2) / gesamt, f"Anbieter {i + 1} von {len(Scans)}: Ergebnisse werden geprüft")

            try:
                agentoutput = run_agent(SYSTEM_PROMPT)
            except Exception as e:
                print(f"Fehler beim Agenten bei Anbieter {name}: {e}")
                failed.append({
                    "anbieter": name,
                    "fehler": str(e)
                })
                agentoutput = ""

            bieterkommentare = get_comments(pydanticobj)

            validation_data = validate_offer()
            final_LV = create_LV(pydanticobj, Soll_LV, validation_data)

            on_progress(0.05 + 0.95 * (i * 4 + 3) / gesamt, f"Anbieter {i + 1} von {len(Scans)}: Excel und Bericht werden erstellt")

            if Dateiformat == "Excel":
                write_excel(final_LV, Ausgang, name, Input_LV)
            elif Dateiformat == "ONLV":
                create_angebotsLV(final_LV, Ausgang, name, tree)
            else:
                raise ValueError("Falsches Dateiformat")

            pdf_report.create_validation_pdf(
                validation_json=validation_data,
                ai_report=agentoutput,
                Extraktionskommentare= pydanticobj.Bericht,
                output_path= Ausgang / f"{name}_bericht.pdf",
                position_count=len(Soll_LV),
                Bieterkommentare=bieterkommentare,
                history=history
            )
        except Exception as e:
            print(f"Fehler bei Anbieter {name}: {e}")
            failed.append({
                "anbieter": name,
                "fehler": str(e)
            })
        finally:
            delete_images(docs, cropped_images)
            rescan_count.clear()
            history = []

    if failed:
        successful = len(Scans) - len(failed)

        on_progress(
            1.0,
            f"Verarbeitung abgeschlossen: {successful}/{len(Scans)} Anbieter erfolgreich. "
            f"{len(failed)} fehlgeschlagen."
        )
        for error in failed:
            print(f"Fehler bei {error['anbieter']}: {error['fehler']}")

    else:
        on_progress(1.0, "Alle Anbieter erfolgreich verarbeitet.")
    return failed

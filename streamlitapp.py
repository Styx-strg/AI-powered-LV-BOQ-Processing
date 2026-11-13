import time
import traceback
from html import escape

import streamlit as st
from datetime import datetime
from pathlib import Path


import Workflow

# erlaubte Endungen für das Soll_LV
SOLL_TYPEN = {
    "ONLV": ["onlv", "xml"],
    "Excel": ["xlsx", "xlsm"],
}
PROJEKT_ORDNER = Path(__file__).parent / "projekte"

# Dateityp, den der Browser beim Download bekommt
MIME_TYPEN = {
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".xls": "application/vnd.ms-excel",
    ".md": "text/markdown",
    ".txt": "text/plain",
    ".pdf": "application/pdf",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".json": "application/json",
    ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
    ".onlv": "application/xml",
}


def save_uploads(soll_name, soll_bytes, pdf_dateien):
    ordner = PROJEKT_ORDNER / datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    eingang = ordner / "eingang"
    ausgang = ordner / "ausgang"
    eingang.mkdir(parents=True)
    ausgang.mkdir()

    # Soll_LV speichern
    soll_pfad = eingang / Path(soll_name).name
    soll_pfad.write_bytes(soll_bytes)

    pdf_pfade = {}
    for name, inhalt in pdf_dateien.items():
        pfad = eingang / Path(name).name
        pfad.write_bytes(inhalt)
        pdf_pfade[name] = pfad

    return soll_pfad, pdf_pfade, ausgang


def get_bytes(ausgang):
    # Liest alle Dateien aus dem Ausgabeordner: {dateiname: Bytes}
    dateien = {}
    for file in sorted(ausgang.iterdir()):
        if file.is_file():
            dateien[file.name] = file.read_bytes()
    return dateien


#Aufruf des Python Codes
def run_pipeline(soll_typ, soll_name, soll_bytes, pdf_dateien, on_progress):
    on_progress(0.02, "Dateien werden gespeichert")
    soll_pfad, pdf_pfade, ausgang = save_uploads(soll_name, soll_bytes, pdf_dateien)

    failed = Workflow.complete_workflow(soll_pfad, soll_typ, pdf_pfade, ausgang, on_progress)

    dateien = get_bytes(ausgang)
    if not dateien:
        # falls irgendwo ein error aufgetreten ist. Ist aber auch die Error Nachricht falls der Workflow über genau 1 Angebot läuft
        raise RuntimeError(f"Der Workflow hat keine Dateien im Ordner {ausgang} erzeugt.")
    return dateien, failed


# Sperre für andere Nutzer wegen globaler Variablen und Ollama
@st.cache_resource
def job_sperre():
    """
    Ein Dictionary, das für ALLE Browser-Sitzungen dasselbe ist.
    So merkt sich die App, ob gerade schon jemand eine Verarbeitung laufen hat.
    (Das lokale Modell schafft realistisch nur einen Job gleichzeitig.)
    """
    return {"laeuft": False}



# CSS Design. Primär AI generiert.
STYLE = """
<style>
/* Streamlit-Menü und Fußzeile ausblenden, wirkt aufgeräumter */
#MainMenu, footer {visibility: hidden;}

/* schmalere, zentrierte Spalte */
.block-container {max-width: 760px; padding-top: 2.5rem;}

/* Kennzahlen als Karten (halbtransparent, passt zu hellem und dunklem Design) */
[data-testid="stMetric"] {
    background: rgba(128, 128, 128, 0.10);
    border-radius: 12px;
    padding: 14px 18px;
}

/* Knöpfe etwas runder */
.stButton > button, .stDownloadButton > button {
    border-radius: 10px;
    font-weight: 600;
    padding: 0.55rem 1.4rem;
}

/* Fortschritts-Ansicht */
.aktuell {font-size: 1.15rem; font-weight: 600; margin: 0.8rem 0 0.4rem 0;}
.protokoll {opacity: 0.65; font-size: 0.92rem; line-height: 1.6;}
</style>
"""


# verschiedene Ansichten je nach Stufe im Workflow
def ansicht_eingabe(sperre):
    with st.container(border=True):
        st.markdown("#### 1 · Soll-LV")
        typ = st.radio(
            "In welchem Format liegt das Soll-LV vor?",
            list(SOLL_TYPEN),
            format_func=lambda t: f"{t}-Datei",
            horizontal=True,
            key="soll_typ",
        )
        soll = st.file_uploader(
            f"Soll-LV hochladen ({typ})",
            type=SOLL_TYPEN[typ],
            key=f"soll_datei_{typ}",
        )

    with st.container(border=True):
        st.markdown("#### 2 · Angebote der Anbieter")
        pdfs = st.file_uploader(
            "PDF-Scans der ausgefüllten Angebote (mehrere möglich)",
            type=["pdf"],
            accept_multiple_files=True,
            key="pdf_dateien",
        )
        if pdfs:
            st.caption(f"{len(pdfs)} Datei(en) ausgewählt")

    if sperre["laeuft"]:
        st.warning("Es läuft gerade eine andere Verarbeitung. Bitte später nochmal versuchen.")

    start_moeglich = bool(soll) and bool(pdfs) and not sperre["laeuft"]
    if not (soll and pdfs) and not sperre["laeuft"]:
        st.caption("Soll-LV und mindestens ein Anbieter-Angebot auswählen.")

    if st.button("▶  Verarbeitung starten", type="primary", disabled=not start_moeglich) and soll is not None:
        # Die Upload-Felder im session state speichern
        st.session_state.job = {
            "soll_typ": typ,
            "soll_name": soll.name,
            "soll_bytes": soll.getvalue(),
            "pdfs": {f.name: f.getvalue() for f in pdfs},
        }
        st.session_state.phase = "laeuft"
        st.rerun()

    with st.expander("Mögliche Fehler"):
        st.write("Wenn die Seite sagt, dass eine andere Verarbeitung läuft, obwohl "
                 "niemand arbeitet, kann die Sperre hier aufgehoben werden.")
        if st.button("Sperre aufheben"):
            sperre["laeuft"] = False
            st.rerun()


def ansicht_laeuft(sperre):
    job = st.session_state.get("job")
    if not job or sperre["laeuft"]:
        st.session_state.phase = "eingabe"
        st.rerun()

    sperre["laeuft"] = True
    t0 = time.time()

    st.subheader("Die Auswertung läuft …")
    st.caption("Dieses Fenster geöffnet lassen. Das kann einige Minuten dauern.")
    balken = st.progress(0.0, text="Starte …")
    aktuell = st.empty()
    verlauf = st.empty()
    protokoll = []

    def on_progress(anteil, text):
        anteil = min(max(anteil, 0.0), 1.0)
        sek = int(time.time() - t0)
        balken.progress(anteil, text=f"{int(anteil * 100)} %  ·  {sek // 60}:{sek % 60:02d} min")
        aktuell.markdown(f"<div class='aktuell'> {escape(text)}</div>", unsafe_allow_html=True)
        if not protokoll or protokoll[-1] != text:
            protokoll.append(text)
        erledigt = [f"✓ {escape(t)}" for t in protokoll[-7:-1]]  # ohne die aktuelle Zeile
        verlauf.markdown(
            "<div class='protokoll'>" + "<br>".join(erledigt) + "</div>",
            unsafe_allow_html=True,
        )

    try:
        ergebnis, failed = run_pipeline(job["soll_typ"], job["soll_name"], job["soll_bytes"], job["pdfs"], on_progress)
        st.session_state.ergebnis = ergebnis
        st.session_state.failed = failed
        st.session_state.phase = "ergebnis"
    except Exception:
        st.session_state.fehler = traceback.format_exc()
        st.session_state.phase = "fehler"
    finally:
        sperre["laeuft"] = False  # wird immer ausgeführt, auch bei Fehlern
    st.rerun()


def reihenfolge(name):
    # Excel-Dateien zuerst, danach alles andere. jeweils nach Name sortiert
    return (Path(name).suffix.lower() != ".xlsx", name)


def ansicht_ergebnis():
    dateien = st.session_state.get("ergebnis")
    if dateien is None:
        zurueck_zum_start()
        st.rerun()

    st.success("Die Auswertung ist fertig.")

    failed = st.session_state.get("failed", [])

    if failed:
        st.warning(f"{len(failed)} Anbieter konnten nicht verarbeitet werden.")

        with st.expander("Fehlerdetails anzeigen"):
            for error in failed:
                st.error(f"**{error['anbieter']}**\n\n{error['fehler']}")

    # Download knopf für jeden Anbieter
    st.markdown("#### Downloads")
    for name in sorted(dateien, key=reihenfolge):
        endung = Path(name).suffix.lower()
        st.download_button(
            f"⬇  {name}",
            data=dateien[name],
            file_name=name,
            mime=MIME_TYPEN.get(endung, "application/octet-stream"),
            type="primary" if endung == ".xlsx" else "secondary",
            key=f"download_{name}",  # jeder Knopf braucht einen eigenen key
        )

    # Berichte als Text zusätzlich direkt anzeigen. bis jetzt unnötig
    for name in sorted(dateien, key=reihenfolge):
        endung = Path(name).suffix.lower()
        if endung in (".md", ".txt"):
            text = dateien[name].decode("utf-8", errors="replace")
            with st.expander(f"Vorschau: {name}", expanded=True):
                if endung == ".md":
                    st.markdown(text)
                else:
                    st.text(text)

    st.write("")
    st.button("Neue Auswertung starten", on_click=zurueck_zum_start)


def ansicht_fehler():
    st.error("Bei der Verarbeitung ist ein Fehler aufgetreten. ")
    with st.expander("Technische Details/Genaue Fehlermeldung"):
        st.code(st.session_state.get("fehler", "keine Details vorhanden"))

    def nochmal():
        st.session_state.phase = "laeuft"

    k1, k2 = st.columns(2)
    k1.button("Nochmal versuchen", type="primary", on_click=nochmal)
    k2.button("Zurück zum Start", on_click=zurueck_zum_start)


def zurueck_zum_start():
    for schluessel in ("job", "ergebnis", "fehler", "failed"):
        st.session_state.pop(schluessel, None)
    st.session_state.phase = "eingabe"



def main():
    st.set_page_config(page_title="LV-Preisspiegel", page_icon="📑")
    st.markdown(STYLE, unsafe_allow_html=True)

    # Damit streamlit nicht bei jedem Klick den Fortschritt löscht.
    st.session_state.setdefault("phase", "eingabe")
    sperre = job_sperre()

    st.title("📑 LV-Preisspiegel")
    st.caption("Handschriftliche Angebote einlesen und als Preisspiegel auswerten. "
               "Die Verarbeitung läuft lokal auf dem Firmen-PC.")

    platz = st.empty()
    with platz.container():
        phase = st.session_state.phase
        if phase == "laeuft":
            ansicht_laeuft(sperre)
        elif phase == "ergebnis":
            ansicht_ergebnis()
        elif phase == "fehler":
            ansicht_fehler()
        else:
            ansicht_eingabe(sperre)

if __name__ == "__main__":
    main()

"""Oberfläche für namitogl.py. Start: streamlit run app.py (oder starten.bat)."""
import io
import zipfile

import streamlit as st

import namitogl as nt


def als_zip(dateien):
    """Mehrere Dateien (Name -> Bytes) als eine ZIP-Datei im Speicher."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, inhalt in dateien.items():
            z.writestr(name, inhalt)
    return buf.getvalue()


st.set_page_config(page_title='namitogl', layout='wide')

st.title('namitogl')
st.write('Aus dem NaMi-Export wird für jede Stufe eine Liste für die Gruppenleitung: nur Vor- und '
         'Nachname, bei Schnuppermitgliedern das Anmeldedatum. Alles läuft auf diesem Computer, '
         'es werden keine Daten verschickt.')

st.subheader('1. NaMi-Export hochladen')
datei = st.file_uploader('Excel-Export der Mitglieder aus NaMi (.xls oder .xlsx)', type=['xls', 'xlsx'])

if datei is None:
    st.info('Sobald ein Export hochgeladen ist, erscheinen hier die Listen zur Prüfung.')
    st.stop()

try:
    export = nt.lies_export(datei)
except Exception:   # beliebiger Lesefehler, z. B. Datei ist keine echte Excel-Datei
    st.error('Die Datei konnte nicht gelesen werden. Bitte in Excel öffnen, '
             'als .xlsx speichern und diese Datei hochladen.')
    st.stop()

try:
    ergebnis = nt.verarbeite(export)
except KeyError as fehlende_spalte:
    st.error(f'Die Datei sieht nicht wie ein NaMi-Export aus: Spalte {fehlende_spalte} fehlt.')
    st.stop()

# ---------- Schritt 2: prüfen ----------
st.subheader('2. Prüfen')
for spalte, (stufe, liste) in zip(st.columns(len(ergebnis.listen)), ergebnis.listen.items()):
    ueberfaellig = int(liste['Überfällig'].sum())
    spalte.metric(stufe, len(liste))
    if ueberfaellig:
        spalte.caption(f'davon {ueberfaellig} länger als {nt.SCHNUPPER_WARNUNG_WOCHEN} Wochen Schnupper')

reiter = st.tabs(list(ergebnis.listen) + [f'Fehler ({len(ergebnis.fehler)})'])
for tab, (stufe, liste) in zip(reiter, ergebnis.listen.items()):
    with tab:
        if liste.empty:
            st.write('Keine Mitglieder in dieser Stufe.')
            continue
        anzeige = liste.drop(columns=['Überfällig'])
        anzeige['Hinweis'] = liste['Überfällig'].map(
            {True: f'Schnupperzeit über {nt.SCHNUPPER_WARNUNG_WOCHEN} Wochen', False: ''})
        st.dataframe(anzeige.fillna(''), hide_index=True)
with reiter[-1]:
    if ergebnis.fehler.empty:
        st.write('Keine Auffälligkeiten.')
    else:
        st.caption('Diese Einträge in NaMi prüfen. Ohne passende Stufe steht ein Mitglied in keiner Liste.')
        st.dataframe(ergebnis.fehler.fillna(''), hide_index=True)

# ---------- Schritt 3: herunterladen ----------
st.subheader('3. Herunterladen')
dateien = nt.ausgabe_dateien(ergebnis)
st.download_button('Alle Listen als ZIP herunterladen', als_zip(dateien),
                   file_name='stufenlisten.zip', mime='application/zip', type='primary')
st.caption('Die Listen enthalten Namen von Kindern und Jugendlichen. Nur an die jeweilige Gruppenleitung weitergeben.')

for spalte, (name, inhalt) in zip(st.columns(len(dateien)), dateien.items()):
    spalte.download_button(name.removesuffix('.xlsx'), inhalt, file_name=name)

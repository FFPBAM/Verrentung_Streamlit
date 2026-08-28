# Ausschüttungs-VV Rechner

Streamlit-Anwendung für die Vermögensverwaltung der Fürst Fugger Privatbank.
Das Tool simuliert, wie sich unterschiedliche Entnahmehöhen auf verschiedene
Portfoliomischungen auswirken, und wird im Beratungsgespräch eingesetzt.

Kernfrage: Wie lange trägt ein Vermögen eine gewünschte laufende Entnahme, und
mit welcher historischen Wahrscheinlichkeit?

> Historische Ergebnisse sind keine Prognose. Verbindliche Angaben und die
> vollständigen Risikohinweise stehen in der begleitenden Broschüre. Für
> steuerliche Fragen ist ein Steuerberater zuständig.

---

## Inhalt des Repositorys

| Datei | Zweck |
|:---|:---|
| `app.py` | Engine und Streamlit-UI in einer Datei |
| `Daten Verrentung_EUR.xlsx` | Marktdaten, Sheet `Import_Daten` |
| `depot_muster.csv` | Vorlage für den Depot-Upload im Nießbrauchsabschnitt |
| `requirements.txt` | Laufzeitabhängigkeiten |
| `requirements-dev.txt` | nur `pytest`, für die Tests |
| `tests/test_depot.py` | Tests für Depot-Import und Nießbrauchsrechnung |
| `.streamlit/secrets.toml.example` | Vorlage für die Zugangsdaten |
| `.streamlit/config.toml` | Theme in Hausfarbe, Upload-Limit |

Die echte `.streamlit/secrets.toml` ist über `.gitignore` ausgeschlossen und
darf nicht eingecheckt werden.

---

## Lokal starten

```bash
git clone <repo-url>
cd <repo>

python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate

pip install -r requirements.txt

cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# Benutzer und Passwörter in der Datei eintragen

streamlit run app.py
```

## Deployment über Streamlit Cloud

1. Repository verbinden, Hauptdatei `app.py`.
2. Unter **App settings → Secrets** den Block aus
   `.streamlit/secrets.toml.example` eintragen und mit echten Zugangsdaten füllen.
3. `Daten Verrentung_EUR.xlsx` muss im Repository-Wurzelverzeichnis liegen,
   sonst greift nur der Upload-Weg in der Seitenleiste.

## Tests

```bash
pip install -r requirements-dev.txt
pytest
```

Der Import von `app.py` startet weder eine Simulation noch die UI. Die gesamte
Oberfläche liegt in `run_app()`.

---

## Fachlicher Rahmen

**Simulation.** Historische Kohortensimulation, kein Monte Carlo. Über die
Zeitreihe wird ein rollierendes Fenster der Länge `horizon_years × 12` gelegt.
Jede Startmonatskohorte ergibt einen Vermögenspfad. Erfolg heißt: das Vermögen
bleibt über den gesamten Horizont positiv.

**Assets.** MSCI World (NDDUWI, Net Total Return, EUR), REXP, Gold. Die
Differenz zu 100 % wird automatisch als Liquidität mit 0 % Ertrag ergänzt.

**Rechenkette.**
1. Monatliche Nominalrenditen aus Indexständen
2. optionaler Gebührenabzug (monatlich anteilig, multiplikativ)
3. optionale Deflationierung mit deutscher Inflation → Realrenditen
4. gewichtete Portfoliorendite (implizit monatliches Rebalancing)
5. Entnahmesimulation mit Cost-Basis-Steuermodell

**Steuermodell.** Auf jeden Verkauf wird der anteilige Gewinn nach dem
Verhältnis Buchgewinn zu Marktwert ermittelt und besteuert. Die Entnahme ist ein
Nettobetrag, der Bruttoverkauf wird entsprechend hochgerechnet. Nicht
abgebildet: Vorabpauschale, Teilfreistellung, Sparerpauschbetrag, Soli,
Kirchensteuer, Verlustverrechnungstöpfe.

**Entnahme.** Konstanter Betrag pro Monat. Im Realmodus ist die Entnahme damit
inflationsindexiert (konstante Kaufkraft), im Nominalmodus konstant nominal.

---

## Nießbrauch (Tab 4)

Der Tab kennt zwei Wege zum Jahreswert. Beide multiplizieren ihn mit dem nach
§ 14 BewG interpolierten Vervielfältiger und deckeln ihn nach § 16 BewG auf ein
Achtzehnkommasechstel des Vermögens.

**Weg 1 — Renditeannahmen (oberer Abschnitt).**
Jahreswert = gewichtete laufende Ertragsrendite × Startvermögen. Deckelungsbasis
ist das Startvermögen aus der Seitenleiste.

**Weg 2 — Depotdatei (unterer Abschnitt).**
Jahreswert = Summe der konkreten Dividenden und Kupons je Position.
Deckelungsbasis ist die Summe der Kurswerte aus der hochgeladenen Datei. Dieser
Weg ist für die Weitergabe an den Steuerberater gedacht; die Positionsübersicht
lässt sich als CSV exportieren.

### Format der Depot-CSV

Semikolon-getrennt, UTF-8 mit BOM, deutsches Zahlenformat (`1.234,56`).
Spaltennamen in der ersten Zeile. `depot_muster.csv` im Repository ist eine
lauffähige Vorlage, die auch in der App heruntergeladen werden kann.

| Spalte | Pflicht | Inhalt |
|:---|:---|:---|
| `isin` | ja¹ | ISIN, bevorzugt gegenüber WKN |
| `wkn` | ja¹ | WKN als Alternative oder Ergänzung |
| `name` | ja | Wertpapierbezeichnung |
| `assetklasse` | ja | `aktien`, `renten`, `fonds`, `edelmetall`, `liquiditaet`, `sonstige` |
| `bestand` | ja | Stück bei Aktien und Fonds, Nominal bei Renten |
| `waehrung` | ja | ISO-Code, z. B. `EUR`, `USD` |
| `kurs` | ja | Kurs in Handelswährung, bei Renten in Prozent vom Nominal |
| `devisenkurs` | ja | Fremdwährung je 1 EUR, `1` bei EUR |
| `kurswert_eur` | ja | Marktwert in EUR zum Stichtag |
| `ertrag_basis` | ja | `je_stueck`, `prozent_nominal` oder `betrag_gesamt` |
| `ertrag_je_einheit` | ja² | Dividende je Stück bzw. Kupon in Prozent p.a. |
| `ertrag_pa_eur` | ja² | Jahresertrag der Position in EUR |
| `ertragsart` | ja | `dividende`, `kupon`, `ausschuettung`, `thesaurierend`, `kein_ertrag` |
| `zahlungsfrequenz` | nein | `jaehrlich`, `halbjaehrlich`, `quartal`, `monatlich` |
| `faelligkeit` | empfohlen | `JJJJ-MM-TT`, nur bei Renten |
| `ertrag_quelle` | empfohlen | z. B. `prospekt_kupon`, `letzte_ausschuettung`, `schaetzung` |
| `stichtag` | ja | Bewertungsstichtag, `JJJJ-MM-TT`, in jeder Zeile gleich |

¹ mindestens eine der beiden Spalten.
² `ertrag_pa_eur` ist die maßgebliche Größe. Fehlt sie, wird sie aus
`ertrag_basis`, `ertrag_je_einheit`, `bestand` und `devisenkurs` berechnet. Sind
beide vorhanden, wird gerechnet und gegen den gelieferten Wert geprüft; eine
Abweichung über einem Prozent erzeugt eine Warnung.

`ertrag_je_einheit` steht in **Handelswährung**, die Umrechnung läuft über
`devisenkurs`. Die Angabe bleibt dadurch stabil, wenn sich der Devisenkurs ändert.

Die alten Exportspalten (`WP Name`, `Auszahlung in Eur pro Titel`,
`Auszahlung in Eur`, `Aktueller Devisenkurs`) werden automatisch auf das neue
Schema gemappt. Ohne `ertrag_basis` kann der Kupon aber nicht geprüft werden.

### Personenbezogene Daten

Enthält die Datei eine Spalte, deren Name auf Kundenbezug hindeutet
(Depotnummer, Kontonummer, Inhaber, Name, IBAN, Adresse, Geburtsdatum,
Steuernummer, Telefon, E-Mail, Berater), wird der Upload abgelehnt. Diese
Spalten sind vor dem Export zu entfernen.

### Behandlung fälliger Anleihen

Ein Kupon endet mit der Fälligkeit, § 14 BewG verlangt aber einen dauerhaften
Jahreswert. Nach Fälligkeit wird deshalb der Rückzahlungsbetrag (Nominal zu
100 %) zur eingestellten Wiederanlagerendite angelegt. Der Jahreswert ist der
zeitgewichtete Durchschnitt beider Phasen über die Restlebenserwartung T:

```
jahreswert = (min(t_f, T) × ertrag_heute + max(0, T − t_f) × ertrag_wiederanlage) / T
```

Diese Gewichtung ist **undiskontiert**, der Vervielfältiger dagegen ein
Barwertfaktor. Späte Perioden werden dadurch leicht übergewichtet. Die exakte
Alternative wäre eine Aufteilung in einen befristeten und einen aufgeschobenen
Nießbrauch mit getrennten Vervielfältigern. Der Wert ist als Näherung
gekennzeichnet, die gewichtete Restlaufzeit des Rentenanteils wird separat
ausgewiesen.

### Weitere Konventionen

- Laufende Kosten (Vermögensverwaltungs- und Depotgebühr) werden **nicht** vom
  Jahreswert abgezogen. Ob der Nießbraucher sie trägt, richtet sich nach der
  Nießbrauchsvereinbarung. Ein Hinweis steht in der App.
- Thesaurierende Positionen werden mit 0 € angesetzt.
- Die Restlebenserwartung ist eine manuelle Eingabe. Es ist bewusst keine
  Sterbetafel hinterlegt.

---

## Bekannte Baustellen

- `st.set_page_config()` steht in `run_app()` nach `check_login()` und damit
  nach anderen Streamlit-Aufrufen. Streamlit erwartet es als ersten Befehl.
- Der Print-CSS-Block läuft auf Modulebene und damit auch beim reinen Import.
- Die Erfolgskurve rechnet jede Rate über alle Kohorten neu und ist der
  langsamste Teil. Caching wäre möglich.
- Gebühren werden auch auf den Liquiditätsanteil angewendet.
- Monatliches Rebalancing ist eine Modellannahme, keine Abbildung der realen
  Umsetzung im Mandat.
- Bei langem Horizont und kurzer Historie bleiben wenige Kohorten übrig; die
  Erfolgsquote ist dann statistisch schwach.

## Hinweis zur Datenlizenz

`Daten Verrentung_EUR.xlsx` enthält Indexdaten (u. a. NDDUWI, REXP). Vor der
Veröffentlichung in einem **öffentlichen** Repository ist die Weitergabe mit dem
jeweiligen Datenanbieter zu klären. Für ein privates Repository ist das
unkritisch.

## Lizenz

MIT, siehe `LICENSE`. Die Lizenz betrifft den Code, nicht die Marktdaten.

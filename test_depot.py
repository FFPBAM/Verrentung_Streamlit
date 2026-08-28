"""
Tests für den Depot-Import des Nießbrauchsabschnitts.

Ausführen:
    pip install -r requirements-dev.txt
    pytest

Der Import von app.py darf weder eine Simulation noch die Streamlit-UI starten;
genau das prüft test_import_startet_keine_ui() implizit, indem der Import
ohne laufenden Streamlit-Kontext gelingt.
"""

import io
import sys
from pathlib import Path

import numpy as np
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import app  # noqa: E402


# ---------------------------------------------------------------------------
# Zahlenparser
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    "roh, erwartet",
    [
        ("1.234,56", 1234.56),
        ("35.000", 35000.0),
        ("97,87", 97.87),
        ("1,1652", 1.1652),
        ("2,63%", 2.63),
        ("660.36", 660.36),
        ("1 234,50", 1234.50),
        (1500, 1500.0),
    ],
)
def test_to_float(roh, erwartet):
    assert app._to_float(roh) == pytest.approx(erwartet)


def test_to_float_leer():
    for roh in ("", "-", "n/a", None):
        assert not np.isfinite(app._to_float(roh))


# ---------------------------------------------------------------------------
# Einlesen und Validierung
# ---------------------------------------------------------------------------

def test_muster_csv_laedt():
    df = app.load_depot_csv(io.StringIO(app.DEPOT_MUSTER_CSV))
    assert len(df) == 6
    for spalte in app.DEPOT_PFLICHTSPALTEN:
        assert spalte in df.columns


def test_personenbezug_wird_abgelehnt():
    csv = (
        "isin;name;assetklasse;kurswert_eur;ertragsart;Depotnummer;Kundenname\n"
        "X;Y;aktien;100;dividende;123;Mustermann\n"
    )
    with pytest.raises(PermissionError) as exc:
        app.load_depot_csv(io.StringIO(csv))
    assert "Depotnummer" in str(exc.value)
    assert "Kundenname" in str(exc.value)


def test_fehlende_pflichtspalten():
    with pytest.raises(KeyError):
        app.load_depot_csv(io.StringIO("isin;name\nA;B\n"))


def test_alte_spaltennamen_werden_gemappt():
    """Der bisherige Depotexport soll ohne Umbenennung lesbar bleiben."""
    csv = (
        "WKN;WP Name;Bestand;Kurswert;Kurs;Auszahlung in Eur pro Titel;"
        "Auszahlung in Eur;Aktueller  Devisenkurs;Assetklasse;ertragsart\n"
        "DE0008404005;Allianz SE;67;30.183,50;450,5;17,03;1.140,89;1;aktien;dividende\n"
    )
    df = app.load_depot_csv(io.StringIO(csv))
    assert df.loc[0, "name"] == "Allianz SE"
    assert df.loc[0, "kurswert_eur"] == pytest.approx(30183.50)
    assert df.loc[0, "ertrag_pa_eur"] == pytest.approx(1140.89)


# ---------------------------------------------------------------------------
# Ertragsermittlung
# ---------------------------------------------------------------------------

def test_kupon_faktor_zehn_wird_erkannt():
    """Der bekannte Faktor-10-Fehler in der bisherigen Datei muss auffallen."""
    csv = (
        "isin;name;assetklasse;bestand;kurswert_eur;ertrag_basis;"
        "ertrag_je_einheit;ertrag_pa_eur;ertragsart;faelligkeit\n"
        "XS3229496180;Deutsche Post;renten;35.000;35.049,52;prozent_nominal;"
        "3,00;105;kupon;2031-11-25\n"
    )
    df = app.load_depot_csv(io.StringIO(csv))
    daten, warnungen = app.berechne_ertraege(df)
    assert daten.loc[0, "ertrag_pa_eur_calc"] == pytest.approx(1050.0)
    assert any("weicht" in w for w in warnungen)


def test_dividende_in_fremdwaehrung():
    csv = (
        "isin;name;assetklasse;bestand;waehrung;devisenkurs;kurswert_eur;"
        "ertrag_basis;ertrag_je_einheit;ertragsart\n"
        "US00287Y1091;AbbVie;aktien;116;USD;1,1652;25.700,79;je_stueck;6,63;dividende\n"
    )
    df = app.load_depot_csv(io.StringIO(csv))
    daten, _ = app.berechne_ertraege(df)
    assert daten.loc[0, "ertrag_pa_eur_calc"] == pytest.approx(116 * 6.63 / 1.1652)


def test_thesaurierend_ist_null():
    csv = (
        "isin;name;assetklasse;kurswert_eur;ertrag_basis;ertragsart\n"
        "IE00B4L5Y983;Core MSCI World;fonds;50.000;betrag_gesamt;thesaurierend\n"
    )
    df = app.load_depot_csv(io.StringIO(csv))
    daten, warnungen = app.berechne_ertraege(df)
    assert daten.loc[0, "ertrag_pa_eur_calc"] == 0.0
    assert any("thesaurierend" in w for w in warnungen)


# ---------------------------------------------------------------------------
# Nießbrauchswert aus dem Depot
# ---------------------------------------------------------------------------

def test_wiederanlage_zeitgewichtet():
    """
    Kupon 1.050 EUR bis Fälligkeit in t Jahren, danach 35.000 * 3 % = 1.050 EUR.
    Bei identischem Niveau muss der Jahreswert unverändert bleiben.
    """
    csv = (
        "isin;name;assetklasse;bestand;kurswert_eur;ertrag_basis;"
        "ertrag_je_einheit;ertragsart;faelligkeit;stichtag\n"
        "X;Anleihe;renten;35.000;35.000,00;prozent_nominal;3,00;kupon;2031-11-25;2026-08-28\n"
    )
    df = app.load_depot_csv(io.StringIO(csv))
    res = app.compute_niessbrauch_aus_depot(df, restlebenserwartung=20.0, wiederanlagerendite=0.03)
    assert res["jahreswert_roh"] == pytest.approx(1050.0, rel=1e-6)


def test_wiederanlage_senkt_jahreswert():
    """Niedrigere Wiederanlagerendite muss den nachhaltigen Jahreswert senken."""
    csv = (
        "isin;name;assetklasse;bestand;kurswert_eur;ertrag_basis;"
        "ertrag_je_einheit;ertragsart;faelligkeit;stichtag\n"
        "X;Anleihe;renten;35.000;35.000,00;prozent_nominal;4,00;kupon;2031-11-25;2026-08-28\n"
    )
    df = app.load_depot_csv(io.StringIO(csv))
    hoch = app.compute_niessbrauch_aus_depot(df, 20.0, 0.04)["jahreswert_roh"]
    tief = app.compute_niessbrauch_aus_depot(df, 20.0, 0.01)["jahreswert_roh"]
    assert tief < hoch


def test_deckelung_paragraph_16():
    """Ein extrem hoher Ertrag muss auf Depotwert / 18,6 gedeckelt werden."""
    csv = (
        "isin;name;assetklasse;kurswert_eur;ertrag_basis;ertrag_pa_eur;ertragsart\n"
        "X;Hochertrag;aktien;100.000,00;betrag_gesamt;50.000,00;dividende\n"
    )
    df = app.load_depot_csv(io.StringIO(csv))
    res = app.compute_niessbrauch_aus_depot(df, 20.0, 0.03)
    assert res["cap_aktiv"] is True
    assert res["jahreswert"] == pytest.approx(100_000.0 / 18.6)


def test_deckelung_nutzt_depotwert_nicht_startvermoegen():
    csv = (
        "isin;name;assetklasse;kurswert_eur;ertrag_basis;ertrag_pa_eur;ertragsart\n"
        "X;A;aktien;60.000,00;betrag_gesamt;1.000,00;dividende\n"
        "Y;B;aktien;40.000,00;betrag_gesamt;500,00;dividende\n"
    )
    df = app.load_depot_csv(io.StringIO(csv))
    res = app.compute_niessbrauch_aus_depot(df, 20.0, 0.03)
    assert res["vermoegen"] == pytest.approx(100_000.0)
    assert res["jahreswert_cap"] == pytest.approx(100_000.0 / 18.6)


def test_niessbrauchswert_ist_jahreswert_mal_vervielfaeltiger():
    df = app.load_depot_csv(io.StringIO(app.DEPOT_MUSTER_CSV))
    res = app.compute_niessbrauch_aus_depot(df, 20.0, 0.03)
    assert res["niessbrauchswert"] == pytest.approx(
        res["jahreswert"] * res["vervielfaeltiger"]
    )


# ---------------------------------------------------------------------------
# Regression: bestehende Engine unverändert
# ---------------------------------------------------------------------------

def test_compute_niessbrauch_unveraendert():
    res = app.compute_niessbrauch(
        initial_wealth=1_000_000.0,
        weights={"msci_world": 0.60, "rexp": 0.35, "Gold": 0.05},
        dividendenrendite=0.01,
        kuponrendite=0.03,
        goldrendite=0.0,
        restlebenserwartung=20.0,
    )
    assert res["rendite_gewichtet"] == pytest.approx(0.0165)
    assert res["jahreswert_roh"] == pytest.approx(16_500.0)
    assert res["cap_aktiv"] is False
    assert res["vervielfaeltiger"] == pytest.approx(12.1)
    assert res["niessbrauchswert"] == pytest.approx(199_650.0)


def test_vervielfaeltiger_interpolation():
    assert app.get_vervielfaeltiger(20) == pytest.approx(12.1)
    assert app.get_vervielfaeltiger(20.5) == pytest.approx(12.25)
    assert app.get_vervielfaeltiger(0.5) == pytest.approx(0.9)
    assert app.get_vervielfaeltiger(95) == pytest.approx(13.4)

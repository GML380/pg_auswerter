library(dplyr)
library(lubridate)
library(readxl)
library(here)
library(hms)
library(openxlsx)

# Parameter
Jahr <- "2025fin"

# Daten einlesen
raw_dat <- read_excel(here::here("input/gt_aufzeichnung.xlsx"), sheet = Jahr)

# Anzahl Teilnehmer
n_teilnehmer <- nrow(raw_dat)

ergebnisse <- raw_dat %>%
  mutate(
    # Uhrzeit ohne Datum
    Laufzeit_Start = hms::as_hms(Laufzeit_Start),
    Laufzeit_Ende  = hms::as_hms(Laufzeit_Ende),
    # Laufzeit in Minuten
    Gehzeit = as.numeric(Laufzeit_Ende - Laufzeit_Start, units = "mins"),
    # Punkte je Kategorie
    Punkte_Laufzeit   = n_teilnehmer + 1 - rank(Gehzeit, ties.method = "first"),         # kürzer besser
    Punkte_XC         = n_teilnehmer + 1 - rank(-XC_Punkte, ties.method = "first"),     # größer besser
    Punkte_Distanz    = n_teilnehmer + 1 - rank(Distanz_Landepunkt, ties.method = "first"), # kleiner besser
    Punkte_Distanz    = if_else(Distanz_Landepunkt == 10000, 0, Punkte_Distanz),
    Punkte_Selfie     = if_else(Selfie == "T", n_teilnehmer, 0),
    Punkte_Außenlandung = if_else(Außenlandung == "T", -0.5 * n_teilnehmer, 0),
    # Gesamtpunkte inkl. Abzug
    Gesamtpunkte = Punkte_Laufzeit + Punkte_XC + Punkte_Distanz + Punkte_Selfie + Punkte_Außenlandung,
    # Gesamtrang
    Rang_Gesamt = rank(-Gesamtpunkte, ties.method = "first")
  ) %>%
  group_by(Geschlecht) %>%
  mutate(Rang_in_Geschlecht = rank(-Gesamtpunkte, ties.method = "first")) %>%
  ungroup() %>%
  mutate(
    Rang_Maenner = if_else(Geschlecht == "M", Rang_in_Geschlecht, NA_real_),
    Rang_Frauen  = if_else(Geschlecht == "F", Rang_in_Geschlecht, NA_real_)
  ) %>%
  select(-Rang_in_Geschlecht) %>% 
  arrange(Rang_Gesamt)

# Hilfstabelle für Ausgabe (Selfie -> Aufstieg)
format_output <- function(df, rang_col) {
  df %>%
    transmute(
      Name = paste(Vorname, Familienname),
      Gehzeit,
      Distanz = Distanz_Landepunkt,
      XC_Punkte,
      Aufstieg = if_else(Selfie == "T", "Gehen", "Gondel"),
      Gesamtpunkte,
      Rang = !!sym(rang_col)
    ) %>%
    arrange(Rang)   # Platz 1 oben
}

# Tabellen vorbereiten
gesamt_tbl <- format_output(ergebnisse, "Rang_Gesamt")
maenner_tbl <- ergebnisse %>% filter(Geschlecht == "M") %>% format_output("Rang_Maenner")
frauen_tbl  <- ergebnisse %>% filter(Geschlecht == "F") %>% format_output("Rang_Frauen")

# Alle Sheets in eine Liste packen
sheets <- list(
  Übersicht = ergebnisse,   # zeigt auch Punkte_Außenlandung
  Gesamt    = gesamt_tbl,
  Männer    = maenner_tbl,
  Frauen    = frauen_tbl
)

# In eine Excel-Datei schreiben
write.xlsx(
  sheets,
  file = here::here("results", paste0("gt_auswertung_", Jahr, ".xlsx"))
)

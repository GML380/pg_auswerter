library(geosphere)
library(readxl)
library(dplyr)
library(tidyverse)
library(writexl)
library(leaflet)
library(htmlwidgets)
library(leaflet)
library(leaflet.extras)  # Für Marker Clustering


# parameter
Jahr <- 2026


# define eichberg start
# XCcontest and google maps use the same definition
pos_start <- c(14.243345, 47.18117) 



# pos_georg <- c(14.19951, 47.18285)  # Georg's point (long, lat)
# 
# # Calculate the Haversine distance (in meters)
# distance_haversine <- distHaversine(pos_start, pos_georg)

# read input file
raw_dat<-read_excel(here::here("input/gleitw_aufzeichnung_neu.xlsx"), sheet = paste(Jahr))

# calculate distance
res_dat<-raw_dat %>% 
  rowwise() %>% 
  mutate(Distanz = distHaversine(c(long, lat), pos_start) / 1000) %>% 
  mutate(Distanz = round(Distanz, 3)) %>% 
  ungroup()

# creating output results
damen<-res_dat %>% 
  filter(Geschlecht == "W") %>% 
  mutate(Platzierung = rank(desc(Distanz))) %>% 
  arrange(Platzierung) %>% 
  select(Platzierung, Vorname, Familienname, Distanz, Hersteller, Typ)

a_b<-res_dat %>% 
  filter(Geschlecht == "M") %>% 
  filter(EN_Kategorie %in% c("A", "B")) %>% 
  mutate(Platzierung = rank(desc(Distanz))) %>% 
  arrange(Platzierung) %>% 
  select(Platzierung, Vorname, Familienname, Distanz, Hersteller, Typ)

offen<-res_dat %>% 
  filter(Geschlecht == "M") %>% 
  filter(!EN_Kategorie %in% c("A", "B")) %>% 
  mutate(Platzierung = rank(desc(Distanz))) %>% 
  arrange(Platzierung) %>% 
  select(Platzierung, Vorname, Familienname, Distanz, Hersteller, Typ)


output_path <- paste0(here::here("results", paste0(Jahr)), "/GW_Ergebnisse_", Jahr, ".xlsx")

# Check if the file exists and remove it if it does
if (file.exists(output_path)) {
  file.remove(output_path)
}

# # Combine and export results
# data_list <- list("Damen" = damen, "A bis B" = a_b, "Offene Klasse" = offen)
# write_xlsx(data_list, path = output_path)

# combine and export results
data_list <- list("Damen" = damen, "A bis B" = a_b, "Offene Klasse" = offen)
write_xlsx(data_list, path = paste0(here::here("results", paste0(Jahr)), "/GW_Ergebnisse_", Jahr,".xlsx"))

# create overview map
map1<-leaflet(data = res_dat) %>%
  addProviderTiles(providers$Esri.WorldImagery) %>%  # Use Esri Satellite Imagery
  addMarkers(
    ~long, ~lat, popup = ~paste(Vorname, Familienname, "|" ,Hersteller, Typ, "|", Distanz, "km"), label = ~paste(Vorname, Familienname, "|" ,Hersteller, Typ, "|", Distanz, "km")
  ) %>%
  addCircleMarkers(
    ~14.243345, ~47.18117,
    color = "black", fillColor = "red", fillOpacity = 1,
    radius = 10,  # Distinctly larger marker
    popup = ~paste("Start"),
    label = ~paste("Start"),
  ) %>%
  setView(lng = mean(c(res_dat$long, 14.243345)), lat = mean(c(res_dat$lat, 47.18117)), zoom = 14)


map1

# Save the map to an HTML file
saveWidget(map1, file = paste0(here::here("results", paste0(Jahr)), "/GW_Map_", Jahr,".html"), selfcontained = TRUE)


# Save the map to an HTML file - for tighub
if (file.exists(paste0(here::here(), "/index.html"))) {
  file.remove(paste0(here::here(), "/index.html"))
}
saveWidget(map1, file = paste0(here::here(), "/index.html"), selfcontained = TRUE)

# # create jpg of the map
# mapshot(
#   map1,
#   file =  paste0(here::here("results"), "/GW_Map_", Jahr,".png"),
#   selfcontained = FALSE
# )


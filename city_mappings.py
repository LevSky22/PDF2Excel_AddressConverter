import re

borough_to_city = {
    "Ahuntsic": "Montreal",
    "Ahuntsic-Cartierville": "Montreal",
    "Anjou": "Montreal",
    "Cote-des-Neiges-Notre-Dame-de-Grace": "Montreal",
    "Côte-des-Neiges–Notre-Dame-de-Grâce": "Montreal",
    "Côte S Luc": "Côte Saint-Luc",
    "CDN/NDG": "Montreal",
    "Lachine": "Montreal",
    "LaSalle": "Montreal",
    "L'Île-Bizard–Sainte-Geneviève": "Montreal",
    "L'Île Biz/Geneviève": "Montreal",
    "Le Plateau-Mont-Royal": "Montreal",
    "Le Plateau M Royal": "Montreal",
    "Le Sud Ouest": "Montreal",
    "Le Sud-Ouest": "Montreal",
    "Mercier–Hochelaga-Maisonneuve": "Montreal",
    "Montréal-Nord": "Montreal",
    "Outremont": "Montreal",
    "Pierrefonds-Roxboro": "Montreal",
    "Plateau-Mont-Royal": "Montreal",
    "Rivière-des-Prairies–Pointe-aux-Trembles": "Montreal",
    "RDP/PAT": "Montreal",
    "Rosemont–La Petite-Patrie": "Montreal",
    "Rosemont": "Montreal",
    "Saint-Laurent": "Montreal",
    "S Laurent": "Montreal",
    "S Léonard": "Montreal",
    "Saint-Léonard": "Montreal",
    "Sud-Ouest": "Montreal",
    "Verdun": "Montreal",
    "Verdun/Île Soeurs": "Montreal",
    "Ville Marie": "Montreal",
    "Villeray–Saint-Michel–Parc-Extension": "Montreal",
    "Villeray/S Michel": "Montreal",
    "Auteuil": "Laval",
    "Chomedey": "Laval",
    "Duvernay": "Laval",
    "Fabreville": "Laval",
    "Îles Laval": "Laval",
    "Laval des Rapides": "Laval",
    "Laval Ouest": "Laval",
    "Laval sur le Lac": "Laval",
    "Pont Viau": "Laval",
    "S François": "Laval",
    "S Martin": "Laval",
    "S Vincent de Paul": "Laval",
    "S Dorothée": "Laval",
    "S Rose": "Laval",
    "Rosemère": "Laval",
    "Vimont": "Laval"
}

# Add a new dictionary for abbreviated city names
abbreviated_cities = {
    "S JEAN RICHELIEU": "Saint-Jean-Sur-Richelieu",
    "S SOPHIE": "Sainte-Sophie",
    "Hemingford Canton": "Hemingford",
    "N D DU LAUS": "NOTRE-DAME-DU-LAUS",
    "S ADOLPHE D'HOWARD": "SAINT-ADOLPHE-D'HOWARD",
    "S AGATHE DES MONTS": "SAINT-AGATHE-DES-MONTS",
    "S AMABLE": "SAINT-AMABLE",
    "S ANNE DES PLAINES": "SAINTE-ANNE-DES-PLAINES",
    "S CATHERINE": "SAINTE-CATHERINE",
    "S AUGUSTIN": "SAINT-AUGUSTIN",
    "S HUBERT": "SAINT-HUBERT",
    "S JÉRÔME": "SAINT-JEROME",
    "S JULIE": "SAINTE-JULIE",
    "S LAZARE": "SAINT-LAZARE"
}

def expand_abbreviated_city(city):
    # Check if the city is in the abbreviated_cities dictionary
    if city.upper() in abbreviated_cities:
        return abbreviated_cities[city.upper()]
    
    # Handle general cases of "S " or "STE " prefixes
    if city.upper().startswith("S "):
        return "ST-" + city[2:].upper()
    elif city.upper().startswith("STE "):
        return "STE-" + city[4:].upper()
    return city

def get_city_from_borough(borough):
    # First, check if it's in the borough_to_city dictionary
    if borough in borough_to_city:
        return borough_to_city[borough]
    
    # If not, try to expand any abbreviations
    expanded_city = expand_abbreviated_city(borough)
    
    # If the expanded city is different from the input, return it
    if expanded_city != borough:
        return expanded_city
    
    # If no match is found, return the original input
    return borough
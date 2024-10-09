borough_to_city = {
    "Ahuntsic": "Montreal",
    "Ahuntsic-Cartierville": "Montreal",
    "Anjou": "Montreal",
    "Baie d'Urfe": "Montreal",
    "Beaconsfield": "Montreal",
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

def get_city_from_borough(borough):
    return borough_to_city.get(borough, borough)
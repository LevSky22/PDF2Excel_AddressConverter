from unidecode import unidecode

REGION_MAPPING = {
    'MONTREAL': [
        # Core Montreal with variations
        'Montréal', 'Montreal', 'Ville de Montréal', 'City of Montreal',
        
        # Boroughs
        'Ahuntsic-Cartierville',
        'Anjou',
        'Côte-des-Neiges–Notre-Dame-de-Grâce', 'Côte-des-Neiges-NDG', 'CDN-NDG',
        'Le Plateau-Mont-Royal', 'Le Plateau', 'Plateau-Mont-Royal',
        'Le Sud-Ouest', 'South West',
        'L\'Île-Bizard–Sainte-Geneviève', 'Ile-Bizard-Sainte-Genevieve',
        'Mercier–Hochelaga-Maisonneuve', 'MHM',
        'Montréal-Nord', 'Montreal North',
        'Outremont',
        'Pierrefonds-Roxboro',
        'Rivière-des-Prairies–Pointe-aux-Trembles', 'RDP-PAT',
        'Rosemont–La Petite-Patrie', 'RPP',
        'Saint-Laurent', 'St-Laurent',
        'Saint-Léonard', 'St-Leonard',
        'Verdun',
        'Ville-Marie',
        'Villeray–Saint-Michel–Parc-Extension', 'VSMPE',

        # Demerged cities
        'Baie-D\'Urfé', 'Baie-Durfe',
        'Beaconsfield',
        'Côte-Saint-Luc', 'Cote-Saint-Luc',
        'Dollard-des-Ormeaux', 'DDO',
        'Dorval',
        'Hampstead',
        'Kirkland',
        'Mont-Royal', 'Mount Royal', 'TMR',
        'Montréal-Est', 'Montreal East',
        'Montréal-Ouest', 'Montreal West',
        'Pointe-Claire',
        'Sainte-Anne-de-Bellevue',
        'Senneville',
        'Westmount',
        'MTL'
    ],

    'MONTEREGIE': [
        # MRC Acton
        'Acton Vale', 'Béthanie', 'Roxton', 'Roxton Falls', 'Saint-Nazaire-d\'Acton',
        'Saint-Théodore-d\'Acton', 'Sainte-Christine', 'Upton',

        # MRC Beauharnois-Salaberry
        'Beauharnois', 'Saint-Étienne-de-Beauharnois', 'Saint-Louis-de-Gonzague',
        'Saint-Stanislas-de-Kostka', 'Saint-Urbain-Premier', 'Sainte-Martine',
        'Salaberry-de-Valleyfield',

        # MRC Brome-Missisquoi
        'Abercorn', 'Bedford', 'Bolton-Ouest', 'Brigham', 'Brome', 'Bromont',
        'Cowansville', 'Dunham', 'East Farnham', 'Farnham', 'Frelighsburg',
        'Lac-Brome', 'Notre-Dame-de-Stanbridge', 'Pike River', 'Saint-Armand',
        'Saint-Ignace-de-Stanbridge', 'Sainte-Sabine', 'Stanbridge East',
        'Stanbridge Station', 'Sutton',

        # MRC La Haute-Yamaska
        'Granby', 'Roxton Pond', 'Saint-Alphonse-de-Granby', 'Saint-Joachim-de-Shefford',
        'Sainte-Cécile-de-Milton', 'Shefford', 'Warden', 'Waterloo',

        # MRC La Vallée-du-Richelieu
        'Beloeil', 'Carignan', 'Chambly', 'McMasterville', 'Mont-Saint-Hilaire',
        'Otterburn Park', 'Saint-Antoine-sur-Richelieu', 'Saint-Basile-le-Grand',
        'Saint-Charles-sur-Richelieu', 'Saint-Denis-sur-Richelieu', 'Saint-Jean-Baptiste',
        'Saint-Marc-sur-Richelieu', 'Saint-Mathieu-de-Beloeil',

        # MRC Le Haut-Richelieu
        'Saint-Jean-sur-Richelieu', 'Saint-Alexandre', 'Saint-Blaise-sur-Richelieu',
        'Saint-Georges-de-Clarenceville', 'Saint-Paul-de-l\'Île-aux-Noix',
        'Saint-Sébastien', 'Saint-Valentin', 'Sainte-Anne-de-Sabrevois',
        'Sainte-Brigide-d\'Iberville', 'Venise-en-Québec',

        # MRC Les Jardins-de-Napierville
        'Hemmingford', 'Napierville', 'Saint-Bernard-de-Lacolle', 'Saint-Cyprien-de-Napierville',
        'Saint-Édouard', 'Saint-Jacques-le-Mineur', 'Saint-Michel', 'Saint-Patrice-de-Sherrington',
        'Saint-Rémi',

        # MRC Les Maskoutains
        'Saint-Hyacinthe', 'Saint-Barnabé-Sud', 'Saint-Bernard-de-Michaudville',
        'Saint-Damase', 'Saint-Dominique', 'Saint-Hugues', 'Saint-Jude',
        'Saint-Liboire', 'Saint-Louis', 'Saint-Marcel-de-Richelieu', 'Saint-Pie',
        'Saint-Simon', 'Saint-Valérien-de-Milton', 'Sainte-Hélène-de-Bagot',
        'Sainte-Madeleine', 'Sainte-Marie-Madeleine',

        # MRC Marguerite-D'Youville
        'Calixa-Lavallée', 'Contrecoeur', 'Saint-Amable', 'Sainte-Julie',
        'Varennes', 'Verchères',

        # MRC Pierre-De Saurel
        'Massueville', 'Saint-Aimé', 'Saint-David', 'Saint-Gérard-Majella',
        'Saint-Joseph-de-Sorel', 'Saint-Ours', 'Saint-Robert', 'Saint-Roch-de-Richelieu',
        'Sainte-Anne-de-Sorel', 'Sainte-Victoire-de-Sorel', 'Sorel-Tracy',
        'Yamaska',

        # MRC Roussillon
        'Candiac', 'Châteauguay', 'Delson', 'La Prairie', 'Léry', 'Mercier',
        'Saint-Constant', 'Saint-Isidore', 'Saint-Mathieu', 'Saint-Philippe',
        'Sainte-Catherine',

        # MRC Rouville
        'Ange-Gardien', 'Marieville', 'Richelieu', 'Rougemont', 'Saint-Césaire',
        'Saint-Mathias-sur-Richelieu', 'Saint-Paul-d\'Abbotsford', 'Sainte-Angèle-de-Monnoir',

        # MRC Vaudreuil-Soulanges
        'Coteau-du-Lac', 'Hudson', 'Les Cèdres', 'Les Coteaux', 'L\'Île-Cadieux',
        'L\'Île-Perrot', 'Notre-Dame-de-l\'Île-Perrot', 'Pincourt', 'Pointe-des-Cascades',
        'Pointe-Fortune', 'Rigaud', 'Rivière-Beaudette', 'Saint-Clet', 'Saint-Lazare',
        'Saint-Polycarpe', 'Saint-Télesphore', 'Saint-Zotique', 'Sainte-Justine-de-Newton',
        'Sainte-Marthe', 'Terrasse-Vaudreuil', 'Très-Saint-Rédempteur', 'Vaudreuil-Dorion',
        'Vaudreuil-sur-le-Lac',

        # Agglomération de Longueuil
        'Boucherville', 'Brossard', 'Longueuil', 'Saint-Bruno-de-Montarville',
        'Saint-Lambert',
        'VSL'
    ],

    'LAURENTIDES': [
        # MRC Antoine-Labelle
        'Chute-Saint-Philippe', 'Ferme-Neuve', 'Kiamika', 'Lac-des-Écorces',
        'Lac-du-Cerf', 'Lac-Saint-Paul', 'La Macaza', 'L\'Ascension',
        'Mont-Laurier', 'Mont-Saint-Michel', 'Nominingue', 'Notre-Dame-de-Pontmain',
        'Notre-Dame-du-Laus', 'Rivière-Rouge', 'Sainte-Anne-du-Lac', 'Saint-Aimé-du-Lac-des-Îles',

        # MRC Argenteuil
        'Brownsburg-Chatham', 'Gore', 'Grenville', 'Grenville-sur-la-Rouge',
        'Harrington', 'Lachute', 'Mille-Isles', 'Saint-André-d\'Argenteuil',
        'Wentworth',

        # MRC Deux-Montagnes
        'Deux-Montagnes', 'Oka', 'Pointe-Calumet', 'Saint-Eustache',
        'Saint-Joseph-du-Lac', 'Saint-Placide', 'Sainte-Marthe-sur-le-Lac',

        # MRC La Rivière-du-Nord
        'Prévost', 'Saint-Colomban', 'Saint-Hippolyte', 'Saint-Jérôme',
        'Sainte-Sophie',

        # MRC Les Laurentides
        'Amherst', 'Arundel', 'Barkmere', 'Brébeuf', 'Huberdeau', 'Ivry-sur-le-Lac',
        'La Conception', 'La Minerve', 'Labelle', 'Lac-Supérieur', 'Lac-Tremblant-Nord',
        'Lantier', 'Mont-Tremblant', 'Montcalm', 'Saint-Faustin–Lac-Carré',
        'Sainte-Agathe-des-Monts', 'Sainte-Lucie-des-Laurentides', 'Val-David',
        'Val-des-Lacs', 'Val-Morin',

        # MRC Les Pays-d'en-Haut
        'Estérel', 'Lac-des-Seize-Îles', 'Morin-Heights', 'Piedmont',
        'Saint-Adolphe-d\'Howard', 'Sainte-Adèle', 'Sainte-Anne-des-Lacs',
        'Sainte-Marguerite-du-Lac-Masson', 'Saint-Sauveur', 'Wentworth-Nord',

        # MRC Thérèse-De Blainville
        'Blainville', 'Boisbriand', 'Bois-des-Filion', 'Lorraine',
        'Rosemère', 'Sainte-Anne-des-Plaines', 'Sainte-Thérèse',

        # Mirabel
        'Mirabel'
    ],

    'LANAUDIERE': [
        # MRC D'Autray
        'Berthierville', 'La Visitation-de-l\'Île-Dupas', 'Lanoraie', 'Lavaltrie',
        'Mandeville', 'Saint-Barthélemy', 'Saint-Cléophas-de-Brandon', 'Saint-Cuthbert',
        'Saint-Didace', 'Saint-Gabriel', 'Saint-Gabriel-de-Brandon', 'Saint-Ignace-de-Loyola',
        'Saint-Norbert', 'Sainte-Élisabeth', 'Sainte-Geneviève-de-Berthier',

        # MRC Joliette
        'Crabtree', 'Joliette', 'Notre-Dame-de-Lourdes', 'Notre-Dame-des-Prairies',
        'Saint-Ambroise-de-Kildare', 'Saint-Charles-Borromée', 'Saint-Paul',
        'Saint-Pierre', 'Saint-Thomas', 'Sainte-Mélanie',

        # MRC L'Assomption
        'Charlemagne', 'L\'Assomption', 'L\'Épiphanie', 'Repentigny',
        'Saint-Sulpice',

        # MRC Les Moulins
        'Mascouche', 'Terrebonne',

        # MRC Matawinie
        'Chertsey', 'Entrelacs', 'Notre-Dame-de-la-Merci', 'Rawdon',
        'Saint-Alphonse-Rodriguez', 'Saint-Côme', 'Saint-Damien', 'Saint-Donat',
        'Saint-Félix-de-Valois', 'Saint-Jean-de-Matha', 'Saint-Michel-des-Saints',
        'Saint-Zénon', 'Sainte-Béatrix', 'Sainte-Émélie-de-l\'Énergie',
        'Sainte-Marcelline-de-Kildare',

        # MRC Montcalm
        'Saint-Alexis', 'Saint-Calixte', 'Saint-Esprit', 'Saint-Jacques',
        'Saint-Liguori', 'Saint-Lin–Laurentides', 'Saint-Roch-de-l\'Achigan',
        'Saint-Roch-Ouest', 'Sainte-Julienne', 'Sainte-Marie-Salomé'
    ],

    'LAVAL': [
        # City
        'Laval', 'Ville de Laval', 'City of Laval',

        # Districts/Neighborhoods
        'Auteuil',
        'Chomedey',
        'Duvernay', 'Duvernay-Est', 'Duvernay-Ouest',
        'Fabreville',
        'Îles-Laval',
        'Laval-des-Rapides',
        'Laval-Ouest', 'Laval West',
        'Laval-sur-le-Lac',
        'Pont-Viau',
        'Saint-François', 'St-François',
        'Saint-Vincent-de-Paul', 'St-Vincent-de-Paul',
        'Sainte-Dorothée', 'Ste-Dorothée',
        'Sainte-Rose', 'Ste-Rose',
        'Vimont'
    ]
}

# Create reverse mapping including all variations
CITY_TO_REGION = {}
for region, cities in REGION_MAPPING.items():
    for city in cities:
        CITY_TO_REGION[city.upper()] = region
        # Add versions without accents
        unaccented = unidecode(city)
        if unaccented != city:
            CITY_TO_REGION[unaccented.upper()] = region

# Higher level regional grouping
SHORE_MAPPING = {
    'NORTH_SHORE': ['LAURENTIDES', 'LANAUDIERE'],
    'SOUTH_SHORE': ['MONTEREGIE'],
    'MONTREAL': ['MONTREAL'],
    'LAVAL': ['LAVAL'],
    'LONGUEUIL': []
}

# Cities in Longueuil agglomeration
LONGUEUIL_CITIES = {
    'LONGUEUIL', 'VIEUX-LONGUEUIL',
    'BROSSARD',
    'SAINT-LAMBERT', 'ST-LAMBERT',
    'BOUCHERVILLE',
    'SAINT-BRUNO-DE-MONTARVILLE', 'ST-BRUNO', 'SAINT-BRUNO',
    'LA PRAIRIE',
    'GREENFIELD PARK',
    'SAINT-HUBERT', 'ST-HUBERT',
    'LEMOYNE', 'LE MOYNE'
}

def get_shore_region(city):
    """
    Get the shore region for a given city.
    Returns one of: flyer_north_shore, flyer_south_shore, flyer_montreal, flyer_laval, flyer_longueuil, flyer_unknown
    """
    if not city:
        return 'unknown'
    
    city_upper = city.upper().strip()
    
    # Handle accents in lookup
    city_unaccented = unidecode(city_upper)
    
    # First check if it's in Longueuil agglomeration
    if city_upper in LONGUEUIL_CITIES or city_unaccented in LONGUEUIL_CITIES:
        return 'longueuil'
    
    # Get the region (e.g., MONTEREGIE, LANAUDIERE, etc.)
    city_region = CITY_TO_REGION.get(city_upper) or CITY_TO_REGION.get(city_unaccented)
    if not city_region:
        return 'unknown'
    
    # Map region to shore
    if city_region == 'MONTEREGIE':
        return 'south_shore'
    elif city_region in ['LAURENTIDES', 'LANAUDIERE']:
        return 'north_shore'
    elif city_region == 'MONTREAL':
        return 'montreal'
    elif city_region == 'LAVAL':
        return 'laval'
    
    return 'unknown'

# Example usage:
if __name__ == "__main__":
    test_cities = [
        'Terrebonne',    # North Shore
        'Brossard',      # Longueuil
        'Montreal',      # Montreal
        'Laval',         # Laval
        'Longueuil',     # Longueuil
        'Invalid City'   # Unknown
    ]
    
    for city in test_cities:
        print(f"{city}: {get_shore_region(city)}")
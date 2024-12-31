import pandas as pd

df_collectie = pd.read_excel(r'C:\Users\flore.verkest\Documents\documenten\code\digiportal\digiportal\dataportal\static\data\collectie.xlsx')
df_collectie = df_collectie[~df_collectie['onderscheidende_kenmerken'].str.contains(r"LEEG", na=False)]
df_collectie = df_collectie[~df_collectie['alternatief_nummer.soort'].str.contains(r"nieuw", na=False)]
df_collectie_iff = df_collectie[df_collectie["instelling.naam"].str.contains('Flanders', na=False)]
df_collectie_ym = df_collectie[~df_collectie["instelling.naam"].str.contains('Flanders', na=False)]
df_collectie_ym = df_collectie_ym[~df_collectie_ym["instelling.naam"].str.contains('Merghelynck', na=False)]
df_collectie_mm = df_collectie[df_collectie["instelling.naam"].str.contains('Merghelynck', na=False)]

df_thesaurus = pd.read_excel(r'C:\Users\flore.verkest\Documents\documenten\code\digiportal\digiportal\dataportal\static\data\thesaurus.xlsx')

######################################################################################################################################################################################
#################################################################### KWALITEITSCONTROLES COLLECTIE IFF ###############################################################################

#################################################################### identificatie ###############################################################################

# instellingsnaam
def iff_001():
    df_001 = df_collectie_iff[df_collectie_iff["instelling.naam"] != 'In Flanders Fields Museum']
    return df_001

# collectie
def iff_002():
    # objectcategorie bevat lege occurences
    df_002 = df_collectie_iff[~df_collectie_iff['collectie'].isna()]
    df_002_001 = df_002[df_002['collectie'].str.startswith('~', na=False)]
    df_002_002 = df_002[df_002['collectie'].str.endswith('~', na=False)]
    df_002_003 = df_002[df_002['collectie'].str.contains('~~', na=False)]
    frames = [df_002_001, df_002_002, df_002_003]
    df_002_01 = pd.concat(frames)
    return df_002_01

# objectnummer
def iff_003():
    # foutieve start objectnummer
    df_003_01 = df_collectie_iff[~df_collectie_iff['objectnummer'].str.startswith(('IFF', 'BAN', 'TBAN', 'IEPWIE', 'PO_', 'MIMAP', 'LBR'))]

    # foutieve lengte objectnummers
    df_01 = df_collectie_iff[df_collectie_iff['objectnummer'].str.startswith('IFF ')]
    df_01 = df_01[~df_01['objectnummer'].apply(lambda x: len(str(x)) == 10)]
    df_01 = df_01[~df_01['objectnummer'].apply(lambda x: len(str(x)) == 13)]
    df_02 = df_collectie_iff[df_collectie_iff['objectnummer'].str.startswith('IFFD')]
    df_02 = df_02[~df_collectie_iff['objectnummer'].str.startswith('IFFDC')]
    df_02 = df_02[~df_collectie_iff['objectnummer'].str.startswith('IFFDA')]
    df_02 = df_02[~df_02['objectnummer'].apply(lambda x: len(str(x)) == 10)]
    df_02 = df_02[~df_02['objectnummer'].apply(lambda x: len(str(x)) == 12)]
    df_03 = df_collectie_iff[df_collectie_iff['objectnummer'].str.startswith('IFFDA')]
    df_03 = df_03[~df_03['objectnummer'].apply(lambda x: len(str(x)) == 8)]
    df_03 = df_03[~df_03['objectnummer'].apply(lambda x: len(str(x)) == 15)]
    df_04 = df_collectie_iff[df_collectie_iff['objectnummer'].str.startswith('IFFDC')]
    df_04 = df_04[~df_04['objectnummer'].apply(lambda x: len(str(x)) == 8)]
    df_04 = df_04[~df_04['objectnummer'].apply(lambda x: len(str(x)) == 15)]
    df_05 = df_collectie_iff[df_collectie_iff['objectnummer'].str.startswith('IFFF')]
    df_05 = df_05[~df_collectie_iff['objectnummer'].str.startswith('IFFFC')]
    df_05 = df_05[~df_collectie_iff['objectnummer'].str.startswith('IFFFA')]
    df_05 = df_05[~df_05['objectnummer'].apply(lambda x: len(str(x)) == 10)]
    df_05 = df_05[~df_05['objectnummer'].apply(lambda x: len(str(x)) == 13)]
    df_06 = df_collectie_iff[df_collectie_iff['objectnummer'].str.startswith('IFFFA')]
    df_06 = df_06[~df_06['objectnummer'].apply(lambda x: len(str(x)) == 8)]
    df_06 = df_06[~df_06['objectnummer'].apply(lambda x: len(str(x)) == 15)]
    df_06 = df_06[~df_06['objectnummer'].apply(lambda x: len(str(x)) == 18)]
    df_07 = df_collectie_iff[df_collectie_iff['objectnummer'].str.startswith('IFFFC')]
    df_07 = df_07[~df_07['objectnummer'].apply(lambda x: len(str(x)) == 8)]
    df_07 = df_07[~df_07['objectnummer'].apply(lambda x: len(str(x)) == 15)]
    df_07 = df_07[~df_07['objectnummer'].apply(lambda x: len(str(x)) == 18)]
    df_08 = df_collectie_iff[df_collectie_iff['objectnummer'].str.startswith('IFFH')]
    df_08 = df_08[~df_08['objectnummer'].apply(lambda x: len(str(x)) == 11)]
    df_08 = df_08[~df_08['objectnummer'].apply(lambda x: len(str(x)) == 14)]
    df_09 = df_collectie_iff[df_collectie_iff['objectnummer'].str.startswith('IFFGD')]
    df_09 = df_09[~df_09['objectnummer'].apply(lambda x: len(str(x)) == 12)]
    df_09 = df_09[~df_09['objectnummer'].apply(lambda x: len(str(x)) == 15)]
    df_10 = df_collectie_iff[df_collectie_iff['objectnummer'].str.startswith('IFFWII')]
    df_10 = df_10[~df_10['objectnummer'].apply(lambda x: len(str(x)) == 13)]
    df_11 = df_collectie_iff[df_collectie_iff['objectnummer'].str.startswith('IFFDB')]
    df_11 = df_11[~df_11['objectnummer'].apply(lambda x: len(str(x)) == 14)]
    df_11 = df_11[~df_11['objectnummer'].apply(lambda x: len(str(x)) == 11)]
    frames = [df_01, df_02, df_03, df_04, df_05, df_06, df_07, df_08, df_09, df_10, df_11]
    df_003_02 = pd.concat(frames)
    df_003_02 = df_003_02[~df_003_02['objectnummer'].str.startswith(('IFF B_', 'IFF D_', 'IFF GB_', 'IFF F_'))]
    return df_003_01, df_003_02

# objectcategorie
def iff_004():
    # objectcategorie bevat lege occurences
    df_004 = df_collectie_iff[~df_collectie_iff['object_categorie'].isna()]
    df_004_001 = df_004[df_004['object_categorie'].str.startswith('~', na=False)]
    df_004_002 = df_004[df_004['object_categorie'].str.endswith('~', na=False)]
    df_004_003 = df_004[df_004['object_categorie'].str.contains('~~', na=False)]
    frames = [df_004_001, df_004_002, df_004_003]
    df_004_01 = pd.concat(frames)
    return df_004_01

# objectnaam
def iff_005():
    # objectnaam is leeg
    df_005_01 = df_collectie_iff[df_collectie_iff['objectnaam'].isna()]

    # objectnaam start met hoofdletter
    df_005_02 = df_collectie_iff[~df_collectie_iff['objectnaam'].isna()]
    df_005_02 = df_005_02[df_005_02['objectnaam'].str.isupper()]

    # objectnaam bevat lege occurences
    df_005 = df_collectie_iff[~df_collectie_iff['objectnaam'].isna()]
    df_005_001 = df_005[df_005['objectnaam'].str.startswith('~', na=False)]
    df_005_002 = df_005[df_005['objectnaam'].str.endswith('~', na=False)]
    df_005_003 = df_005[df_005['objectnaam'].str.contains('~~', na=False)]
    frames = [df_005_001, df_005_002, df_005_003]
    df_005_03 = pd.concat(frames)

    return df_005_01, df_005_02, df_005_03

# titel
def iff_006():
    # titel is leeg
    df_006_01 = df_collectie_iff[df_collectie_iff['titel'].isna()]

    # foutieve start titel (spatie, kleine letter, ...)
    df_006_002 = df_collectie_iff[~df_collectie_iff['titel'].isna()]
    df_006_002 = df_006_002[df_006_002['titel'].str.startswith(' ', na=False)]
    df_006_003 = df_collectie_iff[~df_collectie_iff['titel'].isna()]
    df_006_003['starttitel'] = df_006_003['titel'].astype(str).str[0]
    df_006_003 = df_006_003[~df_006_003['starttitel'].str.isupper()]
    df_006_003 = df_006_003[~df_006_003['starttitel'].str.isdigit()]
    df_006_003.drop(columns=['starttitel'])
    df_006_003 = df_006_003[~df_006_003['titel'].str.startswith('"', na=False)]
    df_006_003 = df_006_003[~df_006_003['titel'].str.startswith("'s ", na=False)]
    df_006_003 = df_006_003[~df_006_003['titel'].str.startswith("'t ", na=False)]
    df_006_003 = df_006_003[~df_006_003['titel'].str.startswith("à", na=False)]
    frames = [df_006_002, df_006_003]
    df_006_02 = pd.concat(frames)

    # titel eindigt op punt
    df_006_03 = df_collectie_iff[~df_collectie_iff['titel'].isna()]
    df_006_03 = df_006_03[df_006_03['titel'].str.endswith(".", na=False)]

    return df_006_01, df_006_02, df_006_03

# beschrijving

######################################################################################################################################################################################
#################################################################### KWALITEITSCONTROLES COLLECTIE YM ################################################################################

# instellingsnaam
def ym_001():
    df_001 = df_collectie_ym[df_collectie_ym["instelling.naam"] != 'Yper Museum']
    df_001 = df_001[df_001["instelling.naam"] != 'Stedelijk Museum Ieper']
    df_001 = df_001[df_001["instelling.naam"] != 'Museum Godshuis Belle']
    df_001 = df_001[df_001["instelling.naam"] != 'Onderwijsmuseum Ieper']
    return df_001

# collectie

# objectnummer

# objectcategorie

# objectnaam
def ym_005():
    # objectnaam is leeg
    df_005_01 = df_collectie_ym[df_collectie_ym['objectnaam'].isna()]

    # objectnaam start met hoofdletter
    df_005_02 = df_collectie_ym[~df_collectie_ym['objectnaam'].isna()]
    df_005_02 = df_005_02[df_005_02['objectnaam'].str.isupper()]

    # objectnaam bevat lege occurences
    df_005 = df_collectie_ym[~df_collectie_ym['objectnaam'].isna()]
    df_005_001 = df_005[df_005['objectnaam'].str.startswith('~', na=False)]
    df_005_002 = df_005[df_005['objectnaam'].str.endswith('~', na=False)]
    df_005_003 = df_005[df_005['objectnaam'].str.contains('~~', na=False)]
    frames = [df_005_001, df_005_002, df_005_003]
    df_005_03 = pd.concat(frames)

    return df_005_01, df_005_02, df_005_03

# titel
def ym_006():
    # titel is leeg
    df_006_01 = df_collectie_ym[df_collectie_ym['titel'].isna()]

    # foutieve start titel (spatie, kleine letter, ...)
    df_006_002 = df_collectie_ym[~df_collectie_ym['titel'].isna()]
    df_006_002 = df_006_002[df_006_002['titel'].str.startswith(' ', na=False)]
    df_006_003 = df_collectie_ym[~df_collectie_ym['titel'].isna()]
    df_006_003['starttitel'] = df_006_003['titel'].astype(str).str[0]
    df_006_003 = df_006_003[~df_006_003['starttitel'].str.isupper()]
    df_006_003 = df_006_003[~df_006_003['starttitel'].str.isdigit()]
    df_006_003.drop(columns=['starttitel'])
    df_006_003 = df_006_003[~df_006_003['titel'].str.startswith('"', na=False)]
    df_006_003 = df_006_003[~df_006_003['titel'].str.startswith("'s ", na=False)]
    df_006_003 = df_006_003[~df_006_003['titel'].str.startswith("'t ", na=False)]
    df_006_003 = df_006_003[~df_006_003['titel'].str.startswith("à", na=False)]
    frames = [df_006_002, df_006_003]
    df_006_02 = pd.concat(frames)

    # titel eindigt op punt
    df_006_03 = df_collectie_ym[~df_collectie_ym['titel'].isna()]
    df_006_03 = df_006_03[df_006_03['titel'].str.endswith(".", na=False)]

    return df_006_01, df_006_02, df_006_03

# beschrijving


######################################################################################################################################################################################
#################################################################### KWALITEITSCONTROLES COLLECTIE MM ################################################################################

# instellingsnaam
def mm_001():
    df_001 = df_collectie_mm[df_collectie_mm["instelling.naam"] != 'Hotel-Museum Arthur Merghelynck']
    return df_001

# collectie

# objectnummer

# objectcategorie

# objectnaam
def mm_005():
    # objectnaam is leeg
    df_005_01 = df_collectie_mm[df_collectie_mm['objectnaam'].isna()]

    # objectnaam start met hoofdletter
    df_005_02 = df_collectie_mm[~df_collectie_mm['objectnaam'].isna()]
    df_005_02 = df_005_02[df_005_02['objectnaam'].str.isupper()]

    # objectnaam bevat lege occurences
    df_005 = df_collectie_mm[~df_collectie_mm['objectnaam'].isna()]
    df_005_001 = df_005[df_005['objectnaam'].str.startswith('~', na=False)]
    df_005_002 = df_005[df_005['objectnaam'].str.endswith('~', na=False)]
    df_005_003 = df_005[df_005['objectnaam'].str.contains('~~', na=False)]
    frames = [df_005_001, df_005_002, df_005_003]
    df_005_03 = pd.concat(frames)

    return df_005_01, df_005_02, df_005_03

# titel
def mm_006():
    # titel is leeg
    df_006_01 = df_collectie_mm[df_collectie_mm['titel'].isna()]

    # foutieve start titel (spatie, kleine letter, ...)
    df_006_002 = df_collectie_mm[~df_collectie_ym['titel'].isna()]
    df_006_002 = df_006_002[df_006_002['titel'].str.startswith(' ', na=False)]
    df_006_003 = df_collectie_mm[~df_collectie_ym['titel'].isna()]
    df_006_003['starttitel'] = df_006_003['titel'].astype(str).str[0]
    df_006_003 = df_006_003[~df_006_003['starttitel'].str.isupper()]
    df_006_003 = df_006_003[~df_006_003['starttitel'].str.isdigit()]
    df_006_003.drop(columns=['starttitel'])
    df_006_003 = df_006_003[~df_006_003['titel'].str.startswith('"', na=False)]
    df_006_003 = df_006_003[~df_006_003['titel'].str.startswith("'s ", na=False)]
    df_006_003 = df_006_003[~df_006_003['titel'].str.startswith("'t ", na=False)]
    df_006_003 = df_006_003[~df_006_003['titel'].str.startswith("à", na=False)]
    frames = [df_006_002, df_006_003]
    df_006_02 = pd.concat(frames)

    # titel eindigt op punt
    df_006_03 = df_collectie_mm[~df_collectie_mm['titel'].isna()]
    df_006_03 = df_006_03[df_006_03['titel'].str.endswith(".", na=False)]

    return df_006_01, df_006_02, df_006_03

# beschrijving


######################################################################################################################################################################################
#################################################################### KWALITEITSCONTROLES BEELDEN IFF #################################################################################


######################################################################################################################################################################################
#################################################################### KWALITEITSCONTROLES BEELDEN YM ##################################################################################


######################################################################################################################################################################################
#################################################################### KWALITEITSCONTROLES BEELDEN MM ###################################################################################


######################################################################################################################################################################################
#################################################################### KWALITEITSCONTROLES THESAURUS IFF ################################################################################

# term
def t_001():
    # term.soort is leeg
    df_001_01 = df_thesaurus[df_thesaurus['term.soort'].isna()]

    # term.status =/ descriptor, non-descriptor
    df_001_02 = df_thesaurus[df_thesaurus["term.status"] != 'descriptor']
    df_001_02 = df_001_02[df_001_02["term.status"] != 'non-descriptor']

    # term start of eindigt met spatie
    df_001_03 = df_thesaurus[df_thesaurus['term'].str.startswith(' ') | df_thesaurus['term'].str.endswith(' ')]

    return df_001_01, df_001_02, df_001_03

# term bron
def t_002():
    # bron start of eindigt met spatie
    df_002_01 = df_thesaurus[df_thesaurus['bron'].str.startswith(' ') | df_thesaurus['bron'].str.endswith(' ')]

    # term nummer start of eindigt met spatie
    df_002_02 = df_thesaurus[df_thesaurus['term.nummer'].str.startswith(' ') | df_thesaurus['term.nummer'].str.endswith(' ')]

    # status descriptor, maar bron en/of scopenote afwezig
    df_002_03 = df_thesaurus[df_thesaurus["term.status"] == 'descriptor']
    types_to_drop = ['rechten', 'afmeting', 'school / stijl', 'toestand']
    df_002_03 = df_002_03[~df_002_03['term.soort'].isin(types_to_drop)]
    df_002_03 = df_002_03[df_002_03["broader_term"] != 'Ieper']
    df_002_03 = df_002_03[df_002_03['bron'].isna()]

    # bron aanwezig, maar nummer ontbreekt
    df_002_04 = df_thesaurus[~df_thesaurus['bron'].isna()]
    df_002_04 = df_002_04[df_002_04['term.nummer'].isna()]

    # nummer aanwezig, maar bron ontbreekt
    df_002_05 = df_thesaurus[~df_thesaurus['term.nummer'].isna()]
    df_002_05 = df_002_05[df_002_05['bron'].isna()]

    # bron AAT, maar nummer =/ 9 digits
    df_002_06 = df_thesaurus
    df_002_06['term.nummer'] = df_002_06['term.nummer'].fillna('').astype(str)
    df_002_06['bron'] = df_002_06['bron'].fillna('')

    df_002_06['is_valid'] = df_002_06.apply(
        lambda row: all(
            len(number) == 9 and number.isdigit()
            for source, number in zip(row['bron'].split('~'), row['term.nummer'].split('~'))
            if source == "http://vocab.getty.edu/aat/"
        ),
        axis=1
    )
    df_002_06 = df_002_06[~df_002_06['is_valid']].drop(columns=['is_valid'])

    # bron Wikidata, maar nummer start niet met Q
    df_002_07 = df_thesaurus
    df_002_07['term.nummer'] = df_002_07['term.nummer'].fillna('').astype(str)
    df_002_07['bron'] = df_002_07['bron'].fillna('')

    df_002_07['is_valid'] = df_002_07.apply(
        lambda row: all(
            number.startswith('Q')
            for source, number in zip(row['bron'].split('~'), row['term.nummer'].split('~'))
            if source == "https://www.wikidata.org/entity/"
        ),
        axis=1
    )
    df_002_07 = df_002_07[~df_002_07['is_valid']].drop(columns=['is_valid'])

    # bron TGN, maar nummer =/ 7 digits
    df_002_08 = df_thesaurus
    df_002_08['term.nummer'] = df_002_08['term.nummer'].fillna('').astype(str)
    df_002_08['bron'] = df_002_08['bron'].fillna('')

    df_002_08['is_valid'] = df_002_08.apply(
        lambda row: all(
            len(number) == 7 and number.isdigit()
            for source, number in zip(row['bron'].split('~'), row['term.nummer'].split('~'))
            if source == "http://vocab.getty.edu/tgn/"
        ),
        axis=1
    )
    df_002_08 = df_002_08[~df_002_08['is_valid']].drop(columns=['is_valid'])

    # foutieve bron
    valid_sources = ['http://vocab.getty.edu/aat/', 'https://www.wikidata.org/entity/', 'https://iconclass.org/', 'http://vocab.getty.edu/tgn/', 'https://id.erfgoed.net/themas/', 
                     'https://id.erfgoed.net/erfgoedobjecten/', 'https://www.geonames.org/','https://www.middeleeuwsmetaal.be/typology-browser', 'https://namenlijst.org/#/memorials/',
                     'https://www.mot.be/resource/Tool/','https://id.erfgoed.net/thesauri/erfgoedtypes/', 'https://id.erfgoed.net/aanduidingsobjecten/', 'http://rightsstatements.org/vocab/'  ]

    df_002_09 = df_thesaurus[~df_thesaurus['bron'].apply(lambda x: all(source in valid_sources for source in x.split('~')))]
    df_002_09 = df_002_09[df_002_09["bron"] != '']

    # List of terms
    termen = df_thesaurus[df_thesaurus["term.status"] == 'non-descriptor']

    # Kolommen waarin je wilt zoeken
    search_columns = ['objectnaam', 'inhoud.onderwerp', 'associatie.onderwerp']

    # Functie om te controleren of een van de termen voorkomt
    def row_contains_term(row, terms):
        for col in search_columns:
            cell_terms = str(row[col]).split('~')  # Splits termen in de cel
            if any(term in cell_terms for term in terms):  # Controleer overlap
                return True
        return False

    # Filter het tweede dataframe
    df_002_10 = df_collectie_iff[df_collectie_iff.apply(row_contains_term, axis=1, terms=termen['term'].tolist())]

    return df_002_01, df_002_02, df_002_03, df_002_04, df_002_05, df_002_06, df_002_07, df_002_08, df_002_09, df_002_10



######################################################################################################################################################################################
#################################################################### KWALITEITSCONTROLES THESAURUS YM ################################################################################


######################################################################################################################################################################################
#################################################################### KWALITEITSCONTROLES THESAURUS MM ################################################################################
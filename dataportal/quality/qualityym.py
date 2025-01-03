import pandas as pd

df_collectie = pd.read_excel(r'C:\Users\flore.verkest\Documents\documenten\code\digiportal\digiportal\dataportal\static\data\collectie.xlsx')
df_collectie = df_collectie[~df_collectie['onderscheidende_kenmerken'].str.contains(r"LEEG", na=False)]
df_collectie = df_collectie[~df_collectie['alternatief_nummer.soort'].str.contains(r"nieuw", na=False)]
df_collectie_ym = df_collectie[~df_collectie["instelling.naam"].str.contains('Flanders', na=False)]
df_collectie_ym = df_collectie_ym[~df_collectie_ym["instelling.naam"].str.contains('Merghelynck', na=False)]

df_thesaurus = pd.read_excel(r'C:\Users\flore.verkest\Documents\documenten\code\digiportal\digiportal\dataportal\static\data\thesaurus.xlsx')

df_beeld = pd.read_excel(r'C:\Users\flore.verkest\Documents\documenten\code\digiportal\digiportal\dataportal\static\data\beeld.xlsx')
df_beeld_hr = df_beeld[df_beeld['pad'].str.contains(r"HogeResolutie", na=False)]
df_beeld_lr = df_beeld[df_beeld['pad'].str.contains(r"LageResolutie", na=False)]
df_beeld_raw = df_beeld[df_beeld['pad'].str.contains(r"RAW", na=False)]
df_beeld_ym = df_beeld[df_beeld['pad'].str.contains(r"02_YM", na=False)]

######################################################################################################################################################################################
#################################################################### KWALITEITSCONTROLES COLLECTIE YM ################################################################################

# identificatie
def ym_q001():
    # instellingsnaam
    df_001_01 = df_collectie_ym[df_collectie_ym["instelling.naam"] != 'Yper Museum']
    df_001_01 = df_001_01[df_001_01["instelling.naam"] != 'Stedelijk Museum Ieper']
    df_001_01 = df_001_01[df_001_01["instelling.naam"] != 'Museum Godshuis Belle']
    df_001_01 = df_001_01[df_001_01["instelling.naam"] != 'Onderwijsmuseum Ieper']

    # collectie
    ## collectie bevat lege occurences
    df_002 = df_collectie_ym[~df_collectie_ym['collectie'].isna()]
    df_002_001 = df_002[df_002['collectie'].str.startswith('~', na=False)]
    df_002_002 = df_002[df_002['collectie'].str.endswith('~', na=False)]
    df_002_003 = df_002[df_002['collectie'].str.contains('~~', na=False)]
    frames = [df_002_001, df_002_002, df_002_003]
    df_002_01 = pd.concat(frames)

    # objectnummer
    ## foutieve start objectnummer
    df_003_01 = df_collectie_ym[~df_collectie_ym['objectnummer'].str.startswith(('YM', 'BAN', 'SM', 'OM', 'MGB', 'MSMAP', 'BRUIK', 'TBAN', 'IEPMER'))]

    # objectcategorie
    ## objectcategorie bevat lege occurences
    df_004 = df_collectie_ym[~df_collectie_ym['object_categorie'].isna()]
    df_004_001 = df_004[df_004['object_categorie'].str.startswith('~', na=False)]
    df_004_002 = df_004[df_004['object_categorie'].str.endswith('~', na=False)]
    df_004_003 = df_004[df_004['object_categorie'].str.contains('~~', na=False)]
    frames = [df_004_001, df_004_002, df_004_003]
    df_004_01 = pd.concat(frames)
    
    # objectnaam
    ## objectnaam is leeg
    df_005_01 = df_collectie_ym[df_collectie_ym['objectnaam'].isna()]

    ## objectnaam start met hoofdletter
    df_005_02 = df_collectie_ym[~df_collectie_ym['objectnaam'].isna()]
    df_005_02 = df_005_02[df_005_02['objectnaam'].str.isupper()]

    ## objectnaam bevat lege occurences
    df_005 = df_collectie_ym[~df_collectie_ym['objectnaam'].isna()]
    df_005_001 = df_005[df_005['objectnaam'].str.startswith('~', na=False)]
    df_005_002 = df_005[df_005['objectnaam'].str.endswith('~', na=False)]
    df_005_003 = df_005[df_005['objectnaam'].str.contains('~~', na=False)]
    frames = [df_005_001, df_005_002, df_005_003]
    df_005_03 = pd.concat(frames)

    # titel
    ## titel is leeg
    df_006_01 = df_collectie_ym[df_collectie_ym['titel'].isna()]

    ## foutieve start titel (spatie, kleine letter, ...)
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

    ## titel eindigt op punt/spatie
    df_006_03 = df_collectie_ym[~df_collectie_ym['titel'].isna()]
    df_006_03 = df_006_03[df_006_03['titel'].str.endswith(('.', ' '), na=False)]

    ## titel is langer dan 250 karakters
    df_006_04 = df_collectie_ym[df_collectie_ym['titel'].str.len() > 250]

    return df_001_01, df_002_01, df_003_01, df_004_01, df_005_01, df_005_02, df_005_03, df_006_01, df_006_02, df_006_03, df_006_04

# vervaardiging
def ym_q002():
    # vervaardiging periode is foutief
    excluded_options = ['Eerste Wereldoorlog','17de eeuw','16de eeuw','20ste eeuw','18de eeuw','14de eeuw','15de eeuw','Tweede Wereldoorlog','13de eeuw','19de eeuw','neolithicum','prehistorie','oudheid','middeleeuwen','vroegmoderne tijd','moderne tijd','eigentijdse tijd']
    df_001_01 = df_collectie_ym[df_collectie_ym['vervaardiging.periode'].apply(lambda periode: all(p not in excluded_options for p in str(periode).split('~')))]
    df_001_01 = df_001_01[~df_001_01['vervaardiging.periode'].isna()]

    # vervaardiging datum begin precisie is foutief
    excluded_options = ['na', 'circa', 'vermoedelijk', 'toegeschreven']
    df_001_02 = df_collectie_ym[df_collectie_ym['vervaardiging.datum.begin.prec'].apply(lambda periode: all(p not in excluded_options for p in str(periode).split('~')))]
    df_001_02 = df_001_02[~df_001_02['vervaardiging.datum.begin.prec'].isna()]

    # vervaardiging datum eind precisie is foutief
    excluded_options = ['voor', 'circa', 'vermoedelijk', 'toegeschreven']
    df_001_03 = df_collectie_ym[df_collectie_ym['vervaardiging.datum.eind.prec'].apply(lambda periode: all(p not in excluded_options for p in str(periode).split('~')))]
    df_001_03 = df_001_03[~df_001_03['vervaardiging.datum.eind.prec'].isna()]

    return df_001_01, df_001_02, df_001_03

# fysieke kenmerken
def ym_q003():   
    # 001 materiaal
    # lege occurences materiaal
    df_001_01 = df_collectie_ym[
        df_collectie_ym['materiaal'].str.startswith('~') | 
        df_collectie_ym['materiaal'].str.endswith('~') | 
        df_collectie_ym['materiaal'].str.contains('~~')
    ]    

    # materiaal ontbreekt
    df_001_02 = df_collectie_ym[df_collectie_ym['materiaal'].isna()]

    # 002 techniek
    # lege occurences techniek
    df_002_01 = df_collectie_ym[
        df_collectie_ym['techniek'].str.startswith('~') | 
        df_collectie_ym['techniek'].str.endswith('~') | 
        df_collectie_ym['techniek'].str.contains('~~')
    ]    

    # techniek ontbreekt
    df_002_02 = df_collectie_ym[df_collectie_ym['techniek'].isna()]

    # 003 afmetingen
    # lege occurences afmetingen
    df_003_01 = df_collectie_ym[
        df_collectie_ym['afmeting.eenheid.lref'].str.startswith('~') | 
        df_collectie_ym['afmeting.eenheid.lref'].str.endswith('~') | 
        df_collectie_ym['afmeting.eenheid.lref'].str.contains('~~')
    ] 

    #afmetingen ontbreken
    df_003_02 = df_collectie_ym[~df_collectie_ym['objectnummer'].str.startswith('YMDB')]
    df_003_02 = df_003_02[df_003_02['afmeting.waarde'].isna()]

    return df_001_01, df_001_02, df_002_01, df_002_02, df_003_01, df_003_02

# iconografie & associaties
def ym_q004():
    #iconografie aanwezig maar soort ontbreekt
    df_001_1 = df_collectie_ym[df_collectie_ym['inhoud.onderwerp'].notna() & df_collectie_ym['inhoud.onderwerp.soort'].isna()]
    df_001_2 = df_collectie_ym[
        df_collectie_ym['inhoud.onderwerp.soort'].str.startswith('~') | 
        df_collectie_ym['inhoud.onderwerp.soort'].str.endswith('~') | 
        df_collectie_ym['inhoud.onderwerp.soort'].str.contains('~~')
    ]
    df_001_01 = pd.concat([df_001_1, df_001_2], ignore_index=True)

    #lege occurences iconografie
    df_001_02 = df_collectie_ym[
        df_collectie_ym['inhoud.onderwerp'].str.startswith('~') | 
        df_collectie_ym['inhoud.onderwerp'].str.endswith('~') | 
        df_collectie_ym['inhoud.onderwerp'].str.contains('~~')
    ]

    #dubbele termen bij iconografie
    df_001_03 = df_collectie_ym[df_collectie_ym['inhoud.onderwerp'].apply(lambda x: isinstance(x, str) and "~" in x and len(x.split("~")) != len(set(x.split("~"))))]

    #soort aanwezig maar iconografie ontbreekt
    df_001_1 = df_collectie_ym[df_collectie_ym['inhoud.onderwerp.soort'].notna() & df_collectie_ym['inhoud.onderwerp'].isna()]
    df_001_2 = df_collectie_ym[
        df_collectie_ym['inhoud.onderwerp'].str.startswith('~') | 
        df_collectie_ym['inhoud.onderwerp'].str.endswith('~') | 
        df_collectie_ym['inhoud.onderwerp'].str.contains('~~')
    ]
    df_001_04 = pd.concat([df_001_1, df_001_2], ignore_index=True)


    #associatie aanwezig maar soort ontbreekt
    df_002_1 = df_collectie_ym[df_collectie_ym['associatie.onderwerp'].notna() & df_collectie_ym['associatie.onderwerp.soort'].isna()]
    df_002_2 = df_collectie_ym[
        df_collectie_ym['associatie.onderwerp.soort'].str.startswith('~') | 
        df_collectie_ym['associatie.onderwerp.soort'].str.endswith('~') | 
        df_collectie_ym['associatie.onderwerp.soort'].str.contains('~~')
    ]
    df_002_01 = pd.concat([df_002_1, df_002_2], ignore_index=True)

    #lege occurences associatie
    df_002_02 = df_collectie_ym[
        df_collectie_ym['associatie.onderwerp'].str.startswith('~') | 
        df_collectie_ym['associatie.onderwerp'].str.endswith('~') | 
        df_collectie_ym['associatie.onderwerp'].str.contains('~~')
    ]

    # associatie periode is foutief
    excluded_options = ['Eerste Wereldoorlog','17de eeuw','16de eeuw','20ste eeuw','18de eeuw','14de eeuw','15de eeuw','Tweede Wereldoorlog','13de eeuw','19de eeuw','neolithicum','prehistorie','oudheid','middeleeuwen','vroegmoderne tijd','moderne tijd','eigentijdse tijd']
    df_002_03 = df_collectie_ym[df_collectie_ym['associatie.periode'].apply(lambda periode: all(p not in excluded_options for p in str(periode).split('~')))]
    df_002_03 = df_002_03[~df_002_03['associatie.periode'].isna()]

    #dubbele termen bij associatie
    df_002_04 = df_collectie_ym[df_collectie_ym['associatie.onderwerp'].apply(lambda x: isinstance(x, str) and "~" in x and len(x.split("~")) != len(set(x.split("~"))))]

    #soort aanwezig maar associatie ontbreekt
    df_002_1 = df_collectie_ym[df_collectie_ym['associatie.onderwerp.soort'].notna() & df_collectie_ym['associatie.onderwerp'].isna()]
    df_002_2 = df_collectie_ym[
        df_collectie_ym['associatie.onderwerp'].str.startswith('~') | 
        df_collectie_ym['associatie.onderwerp'].str.endswith('~') | 
        df_collectie_ym['associatie.onderwerp'].str.contains('~~')
    ]
    df_002_05 = pd.concat([df_002_1, df_002_2], ignore_index=True)
    
    return df_001_01, df_001_02, df_001_03, df_001_04, df_002_01, df_002_02, df_002_03, df_002_04, df_002_05

# rechten
def ym_q005():
    #rechten type ontbreekt
    df_001_01 = df_collectie_ym[df_collectie_ym['rechten.type'].isna()]

    #publiek domein zonder uitleg
    df_001_02 = df_collectie_ym[df_collectie_ym['rechten.type'] == 'Publiek Domein']
    df_001_02 = df_001_02[df_001_02['rechten.startdatum'].isna()]
    df_001_02 = df_001_02[df_001_02['rechten.bijzonderheden'].isna()]

    #in copyright zonder einddatum
    df_001_03 = df_collectie_ym[df_collectie_ym['rechten.type'] == 'In Copyright']
    df_001_03 = df_001_02[df_001_02['rechten.einddatum'].isna()]

    #rechten bijzonderheden foutief
    excluded_options = ['publiek domein: anoniem werk', 'publiek domein: gebrek aan originaliteit', 'risicobepaling: meer dan 150 jaar sinds datum creatie', 'risicobepaling: meer dan 150 jaar sinds geboorte vervaardiger']
    df_001_04 = df_collectie_ym[~df_collectie_ym['rechten.bijzonderheden'].isin(excluded_options)]
    df_001_04 = df_001_04[~df_001_04['rechten.bijzonderheden'].isna()]

    return df_001_01, df_001_02, df_001_03, df_001_04

# verwerving
def ym_q006():
    # verwerving methode is foutief
    excluded_options = ['schenking','aankoop','onbekend','bodemvondst','overdracht','erfpacht','ruil','legaat','bruikleen','teruggave', 'permanente bruikleen']
    df_001_01 = df_collectie_ym[df_collectie_ym['verwerving.methode'].apply(lambda periode: all(p not in excluded_options for p in str(periode).split('~')))]
    df_001_01 = df_001_01[~df_001_01['verwerving.methode'].isna()]

    #verwerving ontbreekt
    df_001_02 = df_collectie_ym[df_collectie_ym['verwerving.methode'].isna()]

    return df_001_01, df_001_02

######################################################################################################################################################################################
#################################################################### KWALITEITSCONTROLES THESAURUS YM ################################################################################

# term
def ym_t001():
    # term.soort is leeg
    df_001_01 = df_thesaurus[df_thesaurus['term.soort'].isna()]

    # term.status =/ descriptor, non-descriptor
    df_001_02 = df_thesaurus[df_thesaurus["term.status"] != 'descriptor']
    df_001_02 = df_001_02[df_001_02["term.status"] != 'non-descriptor']

    # term start of eindigt met spatie
    df_001_03 = df_thesaurus[df_thesaurus['term'].str.startswith(' ') | df_thesaurus['term'].str.endswith(' ')]

    return df_001_01, df_001_02, df_001_03

# term bron
def ym_t002():

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
    df_002_10 = df_collectie_ym[df_collectie_ym.apply(row_contains_term, axis=1, terms=termen['term'].tolist())]

    return df_002_01, df_002_02, df_002_03, df_002_04, df_002_05, df_002_06, df_002_07, df_002_08, df_002_09, df_002_10

######################################################################################################################################################################################
#################################################################### KWALITEITSCONTROLES BEELDEN YM #################################################################################

def ym_b001():

    #beelden gevonden op server, te importeren in CMS
    df_001_01 = df_collectie_ym[df_collectie_ym['reproductie.referentie'].isna()]
    df_001_01 = df_001_01[df_001_01['objectnummer'].isin(df_beeld['objectnummer'])]
    df_001_01 = pd.merge(df_001_01, df_beeld, on="objectnummer", how='outer')
    df_001_01 = df_001_01[~df_001_01['instelling.naam'].isna()]
    df_001_01 = df_001_01.drop_duplicates(subset=['objectnummer'])

    #Records op server niet in adlib 
    object_numbers = tuple(df_beeld_ym['objectnummer'])
    df_001_02 = df_beeld_ym[~df_beeld_ym['objectnummer'].str.startswith(object_numbers)]

    #records in CMS zonder beeld op server
    df_002_01 = df_collectie_ym[~df_collectie_ym['objectnummer'].isin(df_beeld['objectnummer'])]
    df_002_01 = df_002_01[~df_002_01['objectnummer'].str.startswith('MIMAP')]

    #Objecten te digitaliseren
    df_002_02 = df_002_01[df_002_01['objectnummer'].str.startswith(('SM ', 'YM ', 'OM ', 'MGB ', 'YMRG', 'BAN'))]

    #Foto's te digitaliseren
    df_002_03 = df_002_01[df_002_01['objectnummer'].str.startswith(('SMF', 'YMF', 'OMF'))]

    #Documenten te digitaliseren
    df_002_04 = df_002_01[df_002_01['objectnummer'].str.startswith(('SMD', 'YMD', 'OMD'))]

    return df_001_01, df_001_02, df_002_01, df_002_02, df_002_03, df_002_04

def ym_b002():
    # records server raw niet in server hr
    df_001_01 = df_beeld_raw[~df_beeld_raw['objectnummer'].isin(df_beeld_hr['objectnummer'])]

    # records server hr of raw niet in server lr 
    df_01 = df_beeld_hr[~df_beeld_hr['objectnummer'].isin(df_beeld_lr['objectnummer'])]
    df_02 = df_beeld_raw[~df_beeld_raw['objectnummer'].isin(df_beeld_lr['objectnummer'])]
    df_001_02 = pd.concat([df_01, df_02])

    # records server lr niet in server hr of raw
    df_001_03 = df_beeld_lr[~df_beeld_lr['objectnummer'].isin(df_beeld_hr['objectnummer'])]
    df_001_03 = df_001_03[~df_001_03['objectnummer'].isin(df_beeld_raw['objectnummer'])]

    # tif in lr
    options = ['TIF' ,'tif', 'tiff'] 
    df_001_04 = df_beeld_lr[df_beeld_lr['extensie'].isin(options)]   

    # dubbele beelden
    df_002_01 = df_beeld[df_beeld.duplicated(['bestandsnaam', 'bestandsgrootte (MB)'], keep=False)]

    # foutieve server mappen
    valid_directories = {
        'BAN', 'BRUIK', 'A:', 'Musea', '02_YM', 'RELAWM', 'SSGI', 'MGB', 'OM',
        'OMF', 'OMFA', 'OMFC', 'SM', 'SMF', 'SFMC', 'SMGP', 'YM',
        'YMD', 'YMDA', 'YMDB', 'YMF', 'HogeResolutie', 'LageResolutie', '3D', 'RAW'
    }
    valid_prefixes = ('OMFA', 'OMFC', 'SMFC', 'YMDA')

    # Normalize backslashes and convert paths to forward slashes
    df_beeld_ym['normalized_path'] = df_beeld_ym['pad'].str.replace('\\', '/', regex=False)

    # Define a function to validate directories
    def is_valid_path(path):
        directories = path.strip('/').split('/')
        for directory in directories:
            if directory not in valid_directories and not directory.startswith(valid_prefixes):
                return False
        return True

    # Apply the validation function
    df_beeld_ym['is_valid'] = df_beeld_ym['normalized_path'].apply(is_valid_path)

    # Filter out invalid paths
    df_002_02 = df_beeld_ym[~df_beeld_ym['is_valid']].drop(columns=['is_valid', 'normalized_path'])
  
    # foutieve bestandsnamen (df_002_03)
    
    return df_001_01, df_001_02, df_001_03, df_001_04, df_002_01, df_002_02
import pandas as pd

df_collectie = pd.read_excel(r'C:\Users\flore.verkest\Documents\documenten\code\digiportal\digiportal\dataportal\static\data\collectie.xlsx')
df_collectie = df_collectie[~df_collectie['onderscheidende_kenmerken'].str.contains(r"LEEG", na=False)]
df_collectie = df_collectie[~df_collectie['alternatief_nummer.soort'].str.contains(r"nieuw", na=False)]
df_collectie_iff = df_collectie[df_collectie["instelling.naam"].str.contains('Flanders', na=False)]

df_thesaurus = pd.read_excel(r'C:\Users\flore.verkest\Documents\documenten\code\digiportal\digiportal\dataportal\static\data\thesaurus.xlsx')

df_beeld = pd.read_excel(r'C:\Users\flore.verkest\Documents\documenten\code\digiportal\digiportal\dataportal\static\data\beeld.xlsx')
df_beeld_hr = df_beeld[df_beeld['pad'].str.contains(r"HogeResolutie", na=False)]
df_beeld_lr = df_beeld[df_beeld['pad'].str.contains(r"LageResolutie", na=False)]
df_beeld_raw = df_beeld[df_beeld['pad'].str.contains(r"RAW", na=False)]
df_beeld_iff = df_beeld[df_beeld['pad'].str.contains(r"01_IFF", na=False)]

######################################################################################################################################################################################
#################################################################### KWALITEITSCONTROLES COLLECTIE IFF ###############################################################################

# identificatie
def iff_q001():
    # instellingsnaam
    df_001_01 = df_collectie_iff[df_collectie_iff["instelling.naam"] != 'In Flanders Fields Museum']

    # collectie
    ## collectie bevat lege occurences
    df_002 = df_collectie_iff[~df_collectie_iff['collectie'].isna()]
    df_002_001 = df_002[df_002['collectie'].str.startswith('~', na=False)]
    df_002_002 = df_002[df_002['collectie'].str.endswith('~', na=False)]
    df_002_003 = df_002[df_002['collectie'].str.contains('~~', na=False)]
    frames = [df_002_001, df_002_002, df_002_003]
    df_002_01 = pd.concat(frames)

    # objectnummer
    ## foutieve start objectnummer
    df_003_01 = df_collectie_iff[~df_collectie_iff['objectnummer'].str.startswith(('IFF', 'BAN', 'TBAN', 'IEPWIE', 'PO_', 'MIMAP', 'LBR'))]

    ## foutieve lengte objectnummers
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

    # objectcategorie
    ## objectcategorie bevat lege occurences
    df_004 = df_collectie_iff[~df_collectie_iff['object_categorie'].isna()]
    df_004_001 = df_004[df_004['object_categorie'].str.startswith('~', na=False)]
    df_004_002 = df_004[df_004['object_categorie'].str.endswith('~', na=False)]
    df_004_003 = df_004[df_004['object_categorie'].str.contains('~~', na=False)]
    frames = [df_004_001, df_004_002, df_004_003]
    df_004_01 = pd.concat(frames)
    
    # objectnaam
    ## objectnaam is leeg
    df_005_01 = df_collectie_iff[df_collectie_iff['objectnaam'].isna()]

    ## objectnaam start met hoofdletter
    df_005_02 = df_collectie_iff[~df_collectie_iff['objectnaam'].isna()]
    df_005_02 = df_005_02[df_005_02['objectnaam'].str.isupper()]

    ## objectnaam bevat lege occurences
    df_005 = df_collectie_iff[~df_collectie_iff['objectnaam'].isna()]
    df_005_001 = df_005[df_005['objectnaam'].str.startswith('~', na=False)]
    df_005_002 = df_005[df_005['objectnaam'].str.endswith('~', na=False)]
    df_005_003 = df_005[df_005['objectnaam'].str.contains('~~', na=False)]
    frames = [df_005_001, df_005_002, df_005_003]
    df_005_03 = pd.concat(frames)


    # titel
    ## titel is leeg
    df_006_01 = df_collectie_iff[df_collectie_iff['titel'].isna()]

    ## foutieve start titel (spatie, kleine letter, ...)
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

    ## titel eindigt op punt
    df_006_03 = df_collectie_iff[~df_collectie_iff['titel'].isna()]
    df_006_03 = df_006_03[df_006_03['titel'].str.endswith(('.', ' '), na=False)]

    ## titel is langer dan 250 karakters
    df_006_04 = df_collectie_iff[df_collectie_iff['titel'].str.len() > 250]

    return df_001_01, df_002_01, df_003_01, df_003_02, df_004_01, df_005_01, df_005_02, df_005_03, df_006_01, df_006_02, df_006_03, df_006_04

# vervaardiging
def iff_q002():
    # vervaardiging periode is foutief
    excluded_options = ['Eerste Wereldoorlog','17de eeuw','16de eeuw','20ste eeuw','18de eeuw','14de eeuw','15de eeuw','Tweede Wereldoorlog','13de eeuw','19de eeuw','neolithicum','prehistorie','oudheid','middeleeuwen','vroegmoderne tijd','moderne tijd','eigentijdse tijd']
    df_001_01 = df_collectie_iff[df_collectie_iff['vervaardiging.periode'].apply(lambda periode: all(p not in excluded_options for p in str(periode).split('~')))]
    df_001_01 = df_001_01[~df_001_01['vervaardiging.periode'].isna()]

    # vervaardiging datum begin precisie is foutief
    excluded_options = ['na', 'circa', 'vermoedelijk', 'toegeschreven']
    df_001_02 = df_collectie_iff[df_collectie_iff['vervaardiging.datum.begin.prec'].apply(lambda periode: all(p not in excluded_options for p in str(periode).split('~')))]
    df_001_02 = df_001_02[~df_001_02['vervaardiging.datum.begin.prec'].isna()]

    # vervaardiging datum eind precisie is foutief
    excluded_options = ['voor', 'circa', 'vermoedelijk', 'toegeschreven']
    df_001_03 = df_collectie_iff[df_collectie_iff['vervaardiging.datum.eind.prec'].apply(lambda periode: all(p not in excluded_options for p in str(periode).split('~')))]
    df_001_03 = df_001_03[~df_001_03['vervaardiging.datum.eind.prec'].isna()]

    return df_001_01, df_001_02, df_001_03

# fysieke kenmerken
def iff_q003():   
    # 001 materiaal
    # lege occurences materiaal
    df_001_01 = df_collectie_iff[
        df_collectie_iff['materiaal'].str.startswith('~') | 
        df_collectie_iff['materiaal'].str.endswith('~') | 
        df_collectie_iff['materiaal'].str.contains('~~')
    ]    

    # materiaal ontbreekt
    df_001_02 = df_collectie_iff[df_collectie_iff['materiaal'].isna()]

    # 002 techniek
    # lege occurences techniek
    df_002_01 = df_collectie_iff[
        df_collectie_iff['techniek'].str.startswith('~') | 
        df_collectie_iff['techniek'].str.endswith('~') | 
        df_collectie_iff['techniek'].str.contains('~~')
    ]    

    # techniek ontbreekt
    df_002_02 = df_collectie_iff[df_collectie_iff['techniek'].isna()]

    # 003 afmetingen
    # lege occurences afmetingen
    df_003_01 = df_collectie_iff[
        df_collectie_iff['afmeting.eenheid'].str.startswith('~') | 
        df_collectie_iff['afmeting.eenheid'].str.endswith('~') | 
        df_collectie_iff['afmeting.eenheid'].str.contains('~~')
    ] 

    #afmetingen ontbreken
    df_003_02 = df_collectie_iff[~df_collectie_iff['objectnummer'].str.startswith('IFFDB')]
    df_003_02 = df_003_02[df_003_02['afmeting.waarde'].isna()]

    return df_001_01, df_001_02, df_002_01, df_002_02, df_003_01, df_003_02

# iconografie & associaties
def iff_q004():
    #iconografie aanwezig maar soort ontbreekt
    df_001_1 = df_collectie_iff[df_collectie_iff['inhoud.onderwerp'].notna() & df_collectie_iff['inhoud.onderwerp.soort'].isna()]
    df_001_2 = df_collectie_iff[
        df_collectie_iff['inhoud.onderwerp.soort'].str.startswith('~') | 
        df_collectie_iff['inhoud.onderwerp.soort'].str.endswith('~') | 
        df_collectie_iff['inhoud.onderwerp.soort'].str.contains('~~')
    ]
    df_001_01 = pd.concat([df_001_1, df_001_2], ignore_index=True)

    #lege occurences iconografie
    df_001_02 = df_collectie_iff[
        df_collectie_iff['inhoud.onderwerp'].str.startswith('~') | 
        df_collectie_iff['inhoud.onderwerp'].str.endswith('~') | 
        df_collectie_iff['inhoud.onderwerp'].str.contains('~~')
    ]

    #dubbele termen bij iconografie
    df_001_03 = df_collectie_iff[df_collectie_iff['inhoud.onderwerp'].apply(lambda x: isinstance(x, str) and "~" in x and len(x.split("~")) != len(set(x.split("~"))))]

    #soort aanwezig maar iconografie ontbreekt
    df_001_1 = df_collectie_iff[df_collectie_iff['inhoud.onderwerp.soort'].notna() & df_collectie_iff['inhoud.onderwerp'].isna()]
    df_001_2 = df_collectie_iff[
        df_collectie_iff['inhoud.onderwerp'].str.startswith('~') | 
        df_collectie_iff['inhoud.onderwerp'].str.endswith('~') | 
        df_collectie_iff['inhoud.onderwerp'].str.contains('~~')
    ]
    df_001_04 = pd.concat([df_001_1, df_001_2], ignore_index=True)

    #associatie aanwezig maar soort ontbreekt
    df_002_1 = df_collectie_iff[df_collectie_iff['associatie.onderwerp'].notna() & df_collectie_iff['associatie.onderwerp.soort'].isna()]
    df_002_2 = df_collectie_iff[
        df_collectie_iff['associatie.onderwerp.soort'].str.startswith('~') | 
        df_collectie_iff['associatie.onderwerp.soort'].str.endswith('~') | 
        df_collectie_iff['associatie.onderwerp.soort'].str.contains('~~')
    ]
    df_002_01 = pd.concat([df_002_1, df_002_2], ignore_index=True)

    #lege occurences associatie
    df_002_02 = df_collectie_iff[
        df_collectie_iff['associatie.onderwerp'].str.startswith('~') | 
        df_collectie_iff['associatie.onderwerp'].str.endswith('~') | 
        df_collectie_iff['associatie.onderwerp'].str.contains('~~')
    ]

    # associatie periode is foutief
    excluded_options = ['Eerste Wereldoorlog','17de eeuw','16de eeuw','20ste eeuw','18de eeuw','14de eeuw','15de eeuw','Tweede Wereldoorlog','13de eeuw','19de eeuw','neolithicum','prehistorie','oudheid','middeleeuwen','vroegmoderne tijd','moderne tijd','eigentijdse tijd']
    df_002_03 = df_collectie_iff[df_collectie_iff['associatie.periode'].apply(lambda periode: all(p not in excluded_options for p in str(periode).split('~')))]
    df_002_03 = df_002_03[~df_002_03['associatie.periode'].isna()]

    #dubbele termen bij associatie
    df_002_04 = df_collectie_iff[df_collectie_iff['associatie.onderwerp'].apply(lambda x: isinstance(x, str) and "~" in x and len(x.split("~")) != len(set(x.split("~"))))]

    #soort aanwezig maar associatie ontbreekt
    df_002_1 = df_collectie_iff[df_collectie_iff['associatie.onderwerp.soort'].notna() & df_collectie_iff['associatie.onderwerp'].isna()]
    df_002_2 = df_collectie_iff[
        df_collectie_iff['associatie.onderwerp'].str.startswith('~') | 
        df_collectie_iff['associatie.onderwerp'].str.endswith('~') | 
        df_collectie_iff['associatie.onderwerp'].str.contains('~~')
    ]
    df_002_05 = pd.concat([df_002_1, df_002_2], ignore_index=True)

    return df_001_01, df_001_02, df_001_03, df_001_04, df_002_01, df_002_02, df_002_03, df_002_04, df_002_05

# rechten
def iff_q005():
    #rechten type ontbreekt
    df_001_01 = df_collectie_iff[df_collectie_iff['rechten.type'].isna()]

    #publiek domein zonder uitleg
    df_001_02 = df_collectie_iff[df_collectie_iff['rechten.type'] == 'Publiek Domein']
    df_001_02 = df_001_02[df_001_02['rechten.startdatum'].isna()]
    df_001_02 = df_001_02[df_001_02['rechten.bijzonderheden'].isna()]

    #in copyright zonder einddatum
    df_001_03 = df_collectie_iff[df_collectie_iff['rechten.type'] == 'In Copyright']
    df_001_03 = df_001_02[df_001_02['rechten.einddatum'].isna()]

    #rechten bijzonderheden foutief
    excluded_options = ['publiek domein: anoniem werk', 'publiek domein: gebrek aan originaliteit', 'risicobepaling: meer dan 150 jaar sinds datum creatie', 'risicobepaling: meer dan 150 jaar sinds geboorte vervaardiger']
    df_001_04 = df_collectie_iff[~df_collectie_iff['rechten.bijzonderheden'].isin(excluded_options)]
    df_001_04 = df_001_04[~df_001_04['rechten.bijzonderheden'].isna()]

    return df_001_01, df_001_02, df_001_03, df_001_04

# verwerving
def iff_q006():
    # verwerving methode is foutief
    excluded_options = ['schenking','aankoop','onbekend','bodemvondst','overdracht','erfpacht','ruil','legaat','bruikleen','teruggave', 'permanente bruikleen']
    df_001_01 = df_collectie_iff[df_collectie_iff['verwerving.methode'].apply(lambda periode: all(p not in excluded_options for p in str(periode).split('~')))]
    df_001_01 = df_001_01[~df_001_01['verwerving.methode'].isna()]

    #verwerving ontbreekt
    df_001_02 = df_collectie_iff[df_collectie_iff['verwerving.methode'].isna()]

    return df_001_01, df_001_02

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
#################################################################### KWALITEITSCONTROLES BEELDEN IFF #################################################################################

def iff_b001():

    #beelden gevonden op server, te importeren in CMS
    df_001_01 = df_collectie_iff[df_collectie_iff['reproductie.referentie'].isna()]
    df_001_01 = df_001_01[df_001_01['objectnummer'].isin(df_beeld['objectnummer'])]
    df_001_01 = pd.merge(df_001_01, df_beeld, on="objectnummer", how='outer')
    df_001_01 = df_001_01[~df_001_01['instelling.naam'].isna()]
    df_001_01 = df_001_01.drop_duplicates(subset=['objectnummer'])

    #Records op server niet in adlib 
    object_numbers = tuple(df_collectie_iff['objectnummer'])
    df_001_02 = df_beeld_iff[~df_beeld_iff['objectnummer'].str.startswith(object_numbers)]

    #records in CMS zonder beeld op server
    df_002_01 = df_collectie_iff[~df_collectie_iff['objectnummer'].isin(df_beeld['objectnummer'])]
    df_002_01 = df_002_01[~df_002_01['objectnummer'].str.startswith('MIMAP')]

    #Objecten te digitaliseren
    df_002_02 = df_002_01[df_002_01['objectnummer'].str.startswith("IFF ")]

    #Foto's te digitaliseren
    df_002_03 = df_002_01[df_002_01['objectnummer'].str.startswith("IFFF")]

    #Documenten te digitaliseren
    df_002_04 = df_002_01[df_002_01['objectnummer'].str.startswith("IFFD")]

    return df_001_01, df_001_02, df_002_01, df_002_02, df_002_03, df_002_04

def iff_b002():
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
        'BAN', 'BRUIK', 'A:', 'Musea', '01_IFF', 'IEP', 'IFF', 'IFFD', 'IFFDA',
        'IFFDB', 'IFFDC', 'IFFF', 'IFFFA', 'IFFFC', 'IFFGD', 'IFFH', 'IFFLF',
        'IFFWII', 'LBR', 'PO', 'HogeResolutie', 'LageResolutie', '3D', 'RAW'
    }
    valid_prefixes = ('IFFDA', 'IFFFA', 'IFFFC')

    # Normalize backslashes and convert paths to forward slashes
    df_beeld_iff['normalized_path'] = df_beeld_iff['pad'].str.replace('\\', '/', regex=False)

    # Define a function to validate directories
    def is_valid_path(path):
        directories = path.strip('/').split('/')
        for directory in directories:
            if directory not in valid_directories and not directory.startswith(valid_prefixes):
                return False
        return True

    # Apply the validation function
    df_beeld_iff['is_valid'] = df_beeld_iff['normalized_path'].apply(is_valid_path)

    # Filter out invalid paths
    df_002_02 = df_beeld_iff[~df_beeld_iff['is_valid']].drop(columns=['is_valid', 'normalized_path'])

    # foutieve bestandsnamen (df_002_03)
    
    return df_001_01, df_001_02, df_001_03, df_001_04, df_002_01, df_002_02

import pandas as pd
import datetime

# nodige velden in export Adlib:
## objectnummer
## instellingsnaam
## objectnaam
## titel
## reproductie.referentie
## toestand

# export Adlib inlezen
df_collectie = pd.read_excel(r'C:\Users\flore.verkest\Documents\Mijn documenten\code\digiportal\digiportal\dataportal\static\data\iffadlib.xlsx')
year = datetime.datetime.now().year

# export Collectiemap T inlezen

# kwaliteitcontroles
## foutieve instellingsnaam
def iff_001():
    df_001 = df_collectie[df_collectie["instelling.naam"] != 'In Flanders Fields Museum']
    aantal = df_001["objectnummer"].count()
    return df_001, aantal

## foutieve start objectnummer
def iff_002():
    df_002 = df_collectie[~df_collectie['objectnummer'].str.startswith(('IFF', 'BAN'))]
    aantal = df_002["objectnummer"].count()
    return df_002, aantal

## foutieve lengte objectnummer
def iff_003():
    df_01 = df_collectie[df_collectie['objectnummer'].str.startswith('IFF ')]
    df_01 = df_01[~df_01['objectnummer'].apply(lambda x: len(str(x)) == 10)]
    df_01 = df_01[~df_01['objectnummer'].apply(lambda x: len(str(x)) == 13)]
    df_02 = df_collectie[df_collectie['objectnummer'].str.startswith('IFFD')]
    df_02 = df_02[~df_collectie['objectnummer'].str.startswith('IFFDC')]
    df_02 = df_02[~df_collectie['objectnummer'].str.startswith('IFFDA')]
    df_02 = df_02[~df_02['objectnummer'].apply(lambda x: len(str(x)) == 10)]
    df_02 = df_02[~df_02['objectnummer'].apply(lambda x: len(str(x)) == 12)]
    df_03 = df_collectie[df_collectie['objectnummer'].str.startswith('IFFDA')]
    df_03 = df_03[~df_03['objectnummer'].apply(lambda x: len(str(x)) == 8)]
    df_03 = df_03[~df_03['objectnummer'].apply(lambda x: len(str(x)) == 15)]
    df_04 = df_collectie[df_collectie['objectnummer'].str.startswith('IFFDC')]
    df_04 = df_04[~df_04['objectnummer'].apply(lambda x: len(str(x)) == 8)]
    df_04 = df_04[~df_04['objectnummer'].apply(lambda x: len(str(x)) == 15)]
    df_05 = df_collectie[df_collectie['objectnummer'].str.startswith('IFFF')]
    df_05 = df_05[~df_collectie['objectnummer'].str.startswith('IFFFC')]
    df_05 = df_05[~df_collectie['objectnummer'].str.startswith('IFFFA')]
    df_05 = df_05[~df_05['objectnummer'].apply(lambda x: len(str(x)) == 10)]
    df_05 = df_05[~df_05['objectnummer'].apply(lambda x: len(str(x)) == 13)]
    df_06 = df_collectie[df_collectie['objectnummer'].str.startswith('IFFFA')]
    df_06 = df_06[~df_06['objectnummer'].apply(lambda x: len(str(x)) == 8)]
    df_06 = df_06[~df_06['objectnummer'].apply(lambda x: len(str(x)) == 15)]
    df_06 = df_06[~df_06['objectnummer'].apply(lambda x: len(str(x)) == 18)]
    df_07 = df_collectie[df_collectie['objectnummer'].str.startswith('IFFFC')]
    df_07 = df_07[~df_07['objectnummer'].apply(lambda x: len(str(x)) == 8)]
    df_07 = df_07[~df_07['objectnummer'].apply(lambda x: len(str(x)) == 15)]
    df_07 = df_07[~df_07['objectnummer'].apply(lambda x: len(str(x)) == 18)]
    df_08 = df_collectie[df_collectie['objectnummer'].str.startswith('IFFH')]
    df_08 = df_08[~df_08['objectnummer'].apply(lambda x: len(str(x)) == 11)]
    df_08 = df_08[~df_08['objectnummer'].apply(lambda x: len(str(x)) == 14)]
    df_09 = df_collectie[df_collectie['objectnummer'].str.startswith('IFFGD')]
    df_09 = df_09[~df_09['objectnummer'].apply(lambda x: len(str(x)) == 12)]
    df_09 = df_09[~df_09['objectnummer'].apply(lambda x: len(str(x)) == 15)]
    df_10 = df_collectie[df_collectie['objectnummer'].str.startswith('IFFWII')]
    df_10 = df_10[~df_10['objectnummer'].apply(lambda x: len(str(x)) == 13)]
    frames = [df_01, df_02, df_03, df_04, df_05, df_06, df_07, df_08, df_09, df_10]
    df_003 = pd.concat(frames)
    aantal = df_003["objectnummer"].count()
    return df_003, aantal

# objectnaam ontbreekt + meest voorkomende objectnamen
def iff_004():
    df_004 = df_collectie[df_collectie['objectnaam'].isna()]
    aantal = df_004['objectnummer'].count()
    df_005 = df_collectie['objectnaam'].str.split('$', expand=True)
    aantal_lengte = len(df_005.columns)

    xs = []
    for i in range(aantal_lengte):
        xs.append(i)

    df_005 = pd.concat([df_005[xs].melt(value_name='objectnaam')])
    df_005.mask(df_005.eq('None')).dropna()
    df_005 = df_005[df_005['objectnaam'].notna()]
    df_005 = df_005['objectnaam'].value_counts()
    df_005 = df_005.head(15)
    return df_004, df_005, aantal

# titel ontbreekt, foutieve titel
def iff_006():
    df_006 = df_collectie[df_collectie['titel'].isna()]
    aantal = df_006['objectnummer'].count()
    df_007 = df_collectie[~df_collectie['titel'].isna()]
    df_007_01 = df_007[df_007['titel'].str.startswith(' ', na=False)]
    df_007_02 = df_007[df_007['titel'].str.endswith(".", na=False)]
    df_007_03 = df_007
    df_007_03['titel'] = df_007['titel'].astype(str).str[0]
    df_007_03 = df_007_03[~df_007_03['titel'].str.isupper()]
    df_007_03 = df_007_03[~df_007_03['titel'].str.isdigit()]
    df_007_03 = df_007_03[~df_007_03['titel'].str.startswith("'s ", na=False)]
    df_007_03 = df_007_03[~df_007_03['titel'].str.startswith("'t ", na=False)]
    search = [r'\)', r'\(', '"']
    df_007_04 = df_007[df_007['titel'].str.contains('|'.join(search), na=False)]
    frames = [df_007_01, df_007_02, df_007_03, df_007_04]
    df_007 = pd.concat(frames)
    aantal2 = df_007['objectnummer'].count()
    return df_006, aantal, df_007, aantal2

#def iff_007():
#    df_007 = df_collectie[df_collectie['reproductie.referentie'].isna()]
#    df_007 = pd.merge(df_007, tschijfdf, on="objectnummer", how="outer")
#    df_007 = df_007[~df_007['instelling.naam'].isna()]
#    return df_007

#afmetingen ontbreken + afmeting niet in correcte eenheid
def iff_009():
    df_009 = df_collectie[df_collectie['afmeting.waarde'].isna()]
    aantal = df_009['objectnummer'].count()
    df_010 = df_collectie['afmeting.eenheid'].str.split('$', expand=True)
    aantal_lengte = len(df_010.columns)

    xs = []
    for i in range(aantal_lengte):
        xs.append(i)

    df_010 = pd.concat([df_010[xs].melt(value_name='afmeting.eenheid')])
    df_010.mask(df_010.eq('None')).dropna()
    df_010 = df_010[df_010['afmeting.eenheid'].notna()]
    df_010 = df_010['afmeting.eenheid'].drop_duplicates()
    df_010 = pd.DataFrame(df_010, columns = ['afmeting.eenheid'])
    df_010 = df_010[(df_010['afmeting.eenheid'] != 'cm') & (df_010['afmeting.eenheid'] != 'mm')]
    if df_010['afmeting.eenheid'].isnull().values.any() == False:
        df_010 = pd.DataFrame()
    else:
        df_010 = df_010
    return df_009, df_010, aantal



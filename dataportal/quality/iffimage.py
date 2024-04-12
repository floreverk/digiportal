import pandas as pd

#alles op T dat start met correct bestandsnaam, eindigt met correcte extensie
df_ofc = pd.read_excel(r'C:\Users\flore.verkest\Documents\Documenten\code\digiportal\digiportal\dataportal\static\data\digidump.xlsx')
column = ['bestandsnaam', 'objectnummer', 'extensie', 'objectnummer_absoluut', 'path', 'filesize (MB)', 'dpi', 'creatie_datum']

#objectnummer, instellingsnaam & reproductiereferentie
df_iff = pd.read_excel(r'C:\Users\flore.verkest\Documents\Documenten\code\digiportal\digiportal\dataportal\static\data\iffadlib.xlsx')

#objectfotocollectie met correct bestandsnaam, eindigt correcte extensie HR & LR
df_ofc_hr = df_ofc[df_ofc['path'].str.contains(r"HogeResolutie", na=False)]
df_ofc_lr = df_ofc[df_ofc['path'].str.contains(r"LageResolutie", na=False)]
df_ofc_raw = df_ofc[df_ofc['path'].str.contains(r"RAW", na=False)]

######################################################## IMAGE CHECK ########################################################

#records in adlib die niet voorkomen in ofc (te digitaliseren)
def iffi_001():
    df_01 = df_iff[~df_iff['objectnummer'].isin(df_ofc['objectnummer'])]
    return df_01

#Objecten te digitaliseren
def iffi_002():
    df_01 = df_iff[~df_iff['objectnummer'].isin(df_ofc['objectnummer'])]
    df_02 = df_01[df_01['objectnummer'].str.startswith("IFF ")]
    return df_02

#Foto's te digitaliseren
def iffi_003():
    df_01 = df_iff[~df_iff['objectnummer'].isin(df_ofc['objectnummer'])]
    df_03 = df_01[df_01['objectnummer'].str.startswith("IFFF")]
    return df_03

#Documenten te digitaliseren
def iffi_004():
    df_01 = df_iff[~df_iff['objectnummer'].isin(df_ofc['objectnummer'])]
    df_04 = df_01[df_01['objectnummer'].str.startswith("IFFD")]
    return df_04

#records in adlib gevonden in ofc
def iffi_005():
    df_05 = df_iff[df_iff['reproductie.referentie'].isna()]
    df_05 = df_05[df_05['objectnummer'].isin(df_ofc['objectnummer'])]
    df_05 = pd.merge(df_05, df_ofc, on="objectnummer", how='outer')
    df_05 = df_05[~df_05['instelling.naam'].isna()]
    return df_05

#records in ofc niet in adlib
def iffi_006():
    df_06 = df_ofc[~df_ofc['objectnummer'].isin(df_iff['objectnummer'])]
    df_06 = df_06[~df_06['objectnummer_absoluut'].isin(df_iff['objectnummer'])]
    df_06 = df_06[~df_06['objectnummer'].str.endswith("V")]
    df_06 = df_06[~df_06['objectnummer'].str.endswith("R")]
    df_06 = df_06[~df_06['objectnummer'].str.endswith("A")]
    df_06 = df_06[~df_06['objectnummer'].str.endswith("B")]
    df_06 = df_06[~df_06['objectnummer'].str.endswith("WF")]
    return df_06

#records in ofc raw niet in ofc hr
def iffi_007():
    df_07 = df_ofc_raw[~df_ofc_raw['objectnummer'].isin(df_ofc_hr['objectnummer'])]
    return df_07

#records in ofc hr of raw niet in ofc lr 
def iffi_008():
    df_08 = df_ofc_hr[~df_ofc_hr['objectnummer'].isin(df_ofc_lr['objectnummer'])]
    df_08_02 = df_ofc_raw[~df_ofc_raw['objectnummer'].isin(df_ofc_lr['objectnummer'])]
    df_08 = pd.concat([df_08, df_08_02])
    return df_08

#records in ofc lr niet in ofc hr of raw
def iffi_009():
    df_09 = df_ofc_lr[~df_ofc_lr['objectnummer'].isin(df_ofc_hr['objectnummer'])]
    df_09 = df_09[~df_09['objectnummer'].isin(df_ofc_raw['objectnummer'])]
    return df_09

#records hr < 600 dpi
def iffi_010():
    df_10 = df_ofc_hr[df_ofc_hr['dpi'] < 300]
    return df_10

#records lr > 72 dpi
def iffi_011():
    df_11 = df_ofc_lr[df_ofc_lr['objectnummer'].isin(df_ofc_hr['objectnummer'])]
    df_11 = df_11[df_11['dpi'] > 72]
    return df_11

# tif in lr
def iffi_012():
    options = ['TIF' ,'tif', 'tiff'] 
    df_12 = df_ofc_lr[df_ofc_lr['extensie'].isin(options)]   
    return df_12

#dubbele beelden
def iffi_013():
    df_13 = df_ofc[df_ofc.duplicated(['bestandsnaam', 'filesize (MB)'], keep=False)]
    return df_13

########################################################STATS########################################################################

#cijfers collectie (aantallen gedigitaliseerd)
dfs_01 = df_iff[df_iff['objectnummer'].isin(df_ofc['objectnummer'])]

aantaliff = dfs_01['objectnummer'].str.startswith('IFF ').sum()
aantaliffda = dfs_01['objectnummer'].str.startswith('IFFDA').sum()
aantaliffdc = dfs_01['objectnummer'].str.startswith('IFFDC').sum()
aantaliffd = dfs_01['objectnummer'].str.startswith('IFFD').sum() - aantaliffda - aantaliffdc
aantalifffa = dfs_01['objectnummer'].str.startswith('IFFFA').sum()
aantalifffc = dfs_01['objectnummer'].str.startswith('IFFFC').sum()
aantalifff = dfs_01['objectnummer'].str.startswith('IFFF').sum() - aantalifffa - aantalifffc
aantaliffh = dfs_01['objectnummer'].str.startswith('IFFH').sum()
aantaliffgd = dfs_01['objectnummer'].str.startswith('IFFGD').sum()
aantaliffwii = dfs_01['objectnummer'].str.startswith('IFFWII').sum()
aantalban = dfs_01['objectnummer'].str.startswith('BAN').sum()
aantalpo = dfs_01['objectnummer'].str.startswith('PO').sum()
aantallbr = dfs_01['objectnummer'].str.startswith('LBR').sum()
aantaliepwie = dfs_01['objectnummer'].str.startswith('IEPWIE').sum()
aantalmimap = dfs_01['objectnummer'].str.startswith('MIMAP').sum()

#cijfers collectie (aantallen te digitaliseren)
dfs_02 = df_iff[~df_iff['objectnummer'].isin(df_ofc['objectnummer'])]

aantaliffafwezig = dfs_02['objectnummer'].str.startswith('IFF ').sum()
aantaliffdaafwezig = dfs_02['objectnummer'].str.startswith('IFFDA').sum()
aantaliffdcafwezig = dfs_02['objectnummer'].str.startswith('IFFDC').sum()
aantaliffdafwezig = dfs_02['objectnummer'].str.startswith('IFFD').sum() - aantaliffda - aantaliffdc
aantalifffaafwezig = dfs_02['objectnummer'].str.startswith('IFFFA').sum()
aantalifffcafwezig = dfs_02['objectnummer'].str.startswith('IFFFC').sum()
aantalifffafwezig = dfs_02['objectnummer'].str.startswith('IFFF').sum() - aantalifffa - aantalifffc
aantaliffhafwezig = dfs_02['objectnummer'].str.startswith('IFFH').sum()
aantaliffgdafwezig = dfs_02['objectnummer'].str.startswith('IFFGD').sum()
aantaliffwiiafwezig = dfs_02['objectnummer'].str.startswith('IFFWII').sum()
aantalbanafwezig = dfs_02['objectnummer'].str.startswith('BAN').sum()
aantalpoafwezig = dfs_02['objectnummer'].str.startswith('PO').sum()
aantallbrafwezig = dfs_02['objectnummer'].str.startswith('LBR').sum()
aantaliepwieafwezig = dfs_02['objectnummer'].str.startswith('IEPWIE').sum()
aantalmimapafwezig = dfs_02['objectnummer'].str.startswith('MIMAP').sum()
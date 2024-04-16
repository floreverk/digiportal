import matplotlib
import pandas as pd
matplotlib.use('Agg')
from matplotlib import pyplot as plt
from io import BytesIO
import base64
from dataportal.quality import iffanalyse
import datetime
import numpy as np

today = datetime.date.today()
year = today.year

df_ofc = pd.read_excel(r'C:\Users\flore.verkest\Documents\Documenten\code\digiportal\digiportal\dataportal\static\data\digidump.xlsx')

def iff_g001():
    aantal_titel = iffanalyse.iff_006()
    aantal_titel = aantal_titel[3]
    aantal_objectnaam = iffanalyse.iff_004()
    aantal_objectnaam = aantal_objectnaam[2]
    aantal_afmeting = iffanalyse.iff_009()
    aantal_afmeting = aantal_afmeting[2]
    aantal_afwezig = [aantal_objectnaam, aantal_titel, aantal_afmeting]
    labels = ['Objectnaam', 'Titel', 'Afmeting']
    
    plt.style.use('seaborn-v0_8-pastel')
    plt.bar(labels, aantal_afwezig)
    plt.title("Ontbrekende basisregistratie")
    bufferg001 = BytesIO()
    plt.savefig(bufferg001, format="png")
    bufferg001.seek(0)
    imageg001_png = bufferg001.getvalue()
    bufferg001.close()
    g001 = base64.b64encode(imageg001_png)
    g001 = g001.decode('utf-8')
    plt.close()
    return(g001)

def iffi_gi001():
    date_1 = df_ofc[df_ofc["datum_creatie"].str.startswith(str(year))]
    date_1 = date_1["objectnummer"].count()
    date_2 = df_ofc[df_ofc["datum_creatie"].str.startswith(str(year-1))]
    date_2 = date_2["objectnummer"].count()
    date_3 = df_ofc[df_ofc["datum_creatie"].str.startswith(str(year-2))]
    date_3 = date_3["objectnummer"].count()
    date_4 = df_ofc[df_ofc["datum_creatie"].str.startswith(str(year-3))]
    date_4 = date_4["objectnummer"].count()
    date_5 = df_ofc[df_ofc["datum_creatie"].str.startswith(str(year-4))]
    date_5 = date_5["objectnummer"].count()
    aantallen = [date_1, date_2, date_3, date_4, date_5]
    datering = [str(year), str(year-1), str(year-2), str(year-3), str(year-4)]
        
    plt.style.use('seaborn-v0_8-pastel')
    plt.bar(datering, aantallen)
    plt.title("Digitalisatie / jaar")
    buffergi001 = BytesIO()
    plt.savefig(buffergi001, format="png")
    buffergi001.seek(0)
    imagegi001_png = buffergi001.getvalue()
    buffergi001.close()
    gi001 = base64.b64encode(imagegi001_png)
    gi001 = gi001.decode('utf-8')
    plt.close()
    return(gi001)

def iffi_gi002():
    dpi72 = df_ofc[df_ofc['dpi']<=72]
    dpi72 = dpi72['objectnummer'].count()
    dpi300 = df_ofc[(df_ofc['dpi']>72) & (df_ofc['dpi']<=300)]
    dpi300 = dpi300['objectnummer'].count()
    dpi600 = df_ofc[(df_ofc['dpi']>300) & (df_ofc['dpi']<=600)]
    dpi600 = dpi600['objectnummer'].count()
    dpi601 = df_ofc[df_ofc['dpi']>600] 
    dpi601 = dpi601['objectnummer'].count()
    dpi = [dpi72, dpi300, dpi600, dpi601]
    labels = ['<=72dpi', '<=300dpi', '<=600dpi', '>600dpi']

    plt.style.use('seaborn-v0_8-pastel')
    plt.bar(labels, dpi)
    plt.title('aantal/resolutie')
    buffergi002 = BytesIO()
    plt.savefig(buffergi002, format="png")
    buffergi002.seek(0)
    imagegi002_png = buffergi002.getvalue()
    buffergi002.close()
    gi002 = base64.b64encode(imagegi002_png)
    gi002 = gi002.decode('utf-8')
    plt.close()
    return(gi002)

def iffi_gi003():
    mb025 = df_ofc[df_ofc['filesize (MB)']<=0.25]
    mb025 = mb025['objectnummer'].count()
    mb05 = df_ofc[(df_ofc['filesize (MB)']>0.25) & (df_ofc['filesize (MB)']<=0.5)]
    mb05 = mb05['objectnummer'].count()
    mb1 = df_ofc[(df_ofc['filesize (MB)']>0.5) & (df_ofc['filesize (MB)']<=1)]
    mb1 = mb1['objectnummer'].count()
    mb5 = df_ofc[(df_ofc['filesize (MB)']>1) & (df_ofc['filesize (MB)']<=5)]
    mb5 = mb5['objectnummer'].count()
    mb25 = df_ofc[(df_ofc['filesize (MB)']>5) & (df_ofc['filesize (MB)']<=25)]
    mb25 = mb25['objectnummer'].count()
    mb50 = df_ofc[(df_ofc['filesize (MB)']>25) & (df_ofc['filesize (MB)']<=50)]
    mb50 = mb50['objectnummer'].count()
    mb100 = df_ofc[(df_ofc['filesize (MB)']>50) & (df_ofc['filesize (MB)']<=100)]
    mb100 = mb100['objectnummer'].count()
    mb500 = df_ofc[(df_ofc['filesize (MB)']>100) & (df_ofc['filesize (MB)']<=500)]
    mb500 = mb500['objectnummer'].count()
    mb1000 = df_ofc[(df_ofc['filesize (MB)']>500) & (df_ofc['filesize (MB)']<=1000)]
    mb1000 = mb1000['objectnummer'].count()
    mb1001 = df_ofc[df_ofc['filesize (MB)']>1000] 
    mb1001 = mb1001['objectnummer'].count()
    mb = [mb025, mb05, mb1, mb5, mb25, mb50, mb100, mb500, mb1000, mb1001]
    labels = ['<=0.25MB', '<=0.5MB', '<=1MB', '<=5MB', '<=25MB', '<=50MB', '<=100MB','<=500MB', '<=1GB', '>1GB']

    plt.style.use('seaborn-v0_8-pastel')
    plt.bar(labels, mb, width=0.5, label=mb)
    plt.title('Aantal/bestandsgrootte')
    plt.xticks(rotation=90)
    plt.tight_layout()
    buffergi003 = BytesIO()
    plt.savefig(buffergi003, format="png")
    buffergi003.seek(0)
    imagegi003_png = buffergi003.getvalue()
    buffergi003.close()
    gi003 = base64.b64encode(imagegi003_png)
    gi003 = gi003.decode('utf-8')
    plt.close()
    return(gi003)

def iffi_gi004():
    date = df_ofc[df_ofc["datum_creatie"].str.startswith(str(year))]
    iff = date[date['objectnummer'].str.startswith('IFF ')]
    iff = iff['objectnummer'].count()
    ifff = date[date['objectnummer'].str.startswith('IFFF')]
    ifff = ifff['objectnummer'].count()
    iffd = date[date['objectnummer'].str.startswith('IFFD')]
    iffd = iffd['objectnummer'].count()
    date2 = df_ofc[df_ofc["datum_creatie"].str.startswith(str(year-1))]
    iff2 = date2[date2['objectnummer'].str.startswith('IFF ')]
    iff2 = iff2['objectnummer'].count()
    ifff2 = date2[date2['objectnummer'].str.startswith('IFFF')]
    ifff2 = ifff2['objectnummer'].count()
    iffd2 = date2[date2['objectnummer'].str.startswith('IFFD')]
    iffd2 = iffd2['objectnummer'].count()
    date3 = df_ofc[df_ofc["datum_creatie"].str.startswith(str(year-2))]
    iff3 = date3[date3['objectnummer'].str.startswith('IFF ')]
    iff3 = iff3['objectnummer'].count()
    ifff3 = date3[date3['objectnummer'].str.startswith('IFFF')]
    ifff3 = ifff3['objectnummer'].count()
    iffd3 = date3[date3['objectnummer'].str.startswith('IFFD')]
    iffd3 = iffd3['objectnummer'].count()
    date4 = df_ofc[df_ofc["datum_creatie"].str.startswith(str(year-3))]
    iff4 = date4[date4['objectnummer'].str.startswith('IFF ')]
    iff4 = iff4['objectnummer'].count()
    ifff4 = date4[date4['objectnummer'].str.startswith('IFFF')]
    ifff4 = ifff4['objectnummer'].count()
    iffd4 = date4[date4['objectnummer'].str.startswith('IFFD')]
    iffd4 = iffd4['objectnummer'].count()
    date5 = df_ofc[df_ofc["datum_creatie"].str.startswith(str(year-4))]
    iff5 = date5[date5['objectnummer'].str.startswith('IFF ')]
    iff5 = iff5['objectnummer'].count()
    ifff5 = date5[date5['objectnummer'].str.startswith('IFFF')]
    ifff5 = ifff5['objectnummer'].count()
    iffd5 = date5[date5['objectnummer'].str.startswith('IFFD')]
    iffd5 = iffd5['objectnummer'].count()
    
    datering = [str(year), str(year-1), str(year-2), str(year-3), str(year-4)]
    iff_aantallen = {
    'iff': np.array([iff, iff2, iff3, iff4, iff5]),
    'ifff': np.array([ifff, ifff2, ifff3, ifff4, ifff5]),
    'iffd': np.array([iffd, iffd2, iffd3, iffd4, iffd5]),
    }
    width = 0.5

    fig, ax = plt.subplots()
    bottom = np.zeros(5)

    for boolean, iff_aantal in iff_aantallen.items():
        p = ax.bar(datering, iff_aantal, width, label=boolean, bottom=bottom)
        bottom += iff_aantal

    ax.set_title("Number of penguins with above average body mass")
    ax.legend(loc="upper right")

    plt.style.use('seaborn-v0_8-pastel')
    plt.title('digitalisatie/medium/jaar')
    buffergi004 = BytesIO()
    plt.savefig(buffergi004, format="png")
    buffergi004.seek(0)
    imagegi004_png = buffergi004.getvalue()
    buffergi004.close()
    gi004 = base64.b64encode(imagegi004_png)
    gi004 = gi004.decode('utf-8')
    plt.close()
    return(gi004)
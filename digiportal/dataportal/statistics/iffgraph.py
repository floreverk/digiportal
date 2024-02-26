import matplotlib
matplotlib.use('Agg')
from matplotlib import pyplot as plt
from io import BytesIO
import base64
from dataportal.quality import iffanalyse

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
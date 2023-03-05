import pandas as pd
import locale
import re

# Set the locale to Turkish (Turkey)
locale.setlocale(locale.LC_ALL, 'tr_TR')

def clear(x):
    non_alpha = r'[^a-zA-ZçğıöşüÇĞİÖŞÜ\u00E7\u011F\u0131\u00F6\u015F\u00FC\u00C7\u011E\u0130\u00D6\u015E\u00DC]'
    ret = x
    ret = re.sub('SN: ','SN:', ret)
    ret = re.sub('SN:\d+','_ ', ret)
    ret = re.sub('Banka: \d+','_ ', ret)
    ret = re.sub('GönBanka:\d+ ','_ ', ret)
    ret = re.sub('GönŞube:\d+ ','_ ', ret)
    ret = re.sub('EftRef:\d+ ','_ ', ret)
    ret = re.sub('Eft Otomatik Muhasebe',' ', ret)
    ret = re.sub('Eft Otomatik M.*?(?=\s|$)','_ ',ret)
    ret = re.sub('Eft O.*?(?=\s|$)','_ ',ret)
    ret = ret.replace('_ ','')
    ret = ret.replace('i','İ').upper()
    """
    ret = ret.replace('Ç','C')
    ret = ret.replace('Ğ','G')
    ret = ret.replace('İ','I')
    ret = ret.replace('Ö','O')
    ret = ret.replace('Ş','S')
    ret = ret.replace('Ü','U')
    """
    ret = re.sub(non_alpha, ' ', ret)
    ret = ret.strip()
    return ret

def cmatch(x, tokens):
    tokens = np.unique(tokens).tolist()
    isim = x.İSİM.split(' ')
    
    anne = clear(x['BABA ADI']).split(' ')
    # anne/baba adı veya soyadı çocuğun ismi içinde olduğu zaman çift sayıyor
    anne = [a for a in anne if not a in isim]
    
    baba = clear(x['ANNE ADI']).split(' ')
    # anne/baba adı veya soyadı çocuğun ismi içinde olduğu zaman çift sayıyor
    baba = [b for b in baba if not b in isim]
    
    m1 = sum([1   for t in tokens for o in isim if t==o])
    m2 = sum([0.3 for t in tokens for o in anne if t==o])
    m3 = sum([0.3 for t in tokens for o in baba if t==o])
    return m1+m2+m3

def aciklama_tokenize(x):
    x = clear(x)
    
    tokens = [t for t in x.split(' ') if t!='']

    dff = dof.assign(PUAN=dof.fillna('').apply(lambda x: cmatch(x, tokens), axis=1))
    dff = dff[dff.PUAN>1].sort_values('PUAN', ascending=False)
    max_puan = dff.PUAN.max()
    dff = dff[dff.PUAN==max_puan]
    
    if dff.shape[0]==0:
        oid=-1
    elif dff.shape[0]==1:
        oid=dff.iloc[0,].name
    else:
        oid=-2
    ret = pd.Series({'Hesap Açıklama':x, 'ÖğrenciID': oid ,'Eşleşenler':dff.reset_index().values})
    
    return ret


df = pd.read_excel('data/Hesap Hareketleri_20220801_20230118.xlsx', skiprows=8)
do = pd.read_excel('data/Bilsem_12_12_2022 14_07_53.xlsx', sheet_name='genel liste')
dfc = df.assign(Açıklama2=df.Açıklama.apply(lambda x: clear(x)))

#dfc.iloc[477:478].apply(lambda x: aciklama_tokenize(x.Açıklama2), axis=1)
dfc.iloc[477:478].assign(lambda x: x.apply(lambda y: aciklama_tokenize(y.Açıklama2), axis=1))

dfdbl = dfc.iloc[dbl].apply(lambda x: aciklama_tokenize(x))
dfdbl

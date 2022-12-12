from string import punctuation, digits
from pandas import read_excel, DataFrame
from numpy import append, delete, where, array
from unidecode import unidecode
from advertools import stopwords as stw
from nltk import Text
from nltk.probability import FreqDist
from nltk.tokenize import word_tokenize

#METODOS
def eliminar_caracteres_especiales(text):
    caracteres_especiales = punctuation + digits + '\xa0\x96\n\t¡¿«»·'
    return "".join([ch for ch in text if ch not in caracteres_especiales]).lower()

def eliminar_tildes(text):
    return unidecode(text)

def eliminar_espacios_inicio_fin(text):
    return text.strip()

def abrir_archivo(ubicacion, sheet):
    return read_excel(ubicacion, sheet_name=sheet)

def obtener_stopwords():
    return [eliminar_tildes(x) for x in list(stw['spanish'])]

def procesar_linea_texto(texto):
    c1 = eliminar_caracteres_especiales(texto)
    c2 = eliminar_tildes(c1)
    return eliminar_espacios_inicio_fin(c2)

def eliminar_palabras(lista):
    for palabra in lista:
        if palabra in stw_spanish:
            lista = delete(lista, where(lista == palabra))

    return lista

#DEFINICIONES
df = abrir_archivo("C:\\Users\\richa\\OneDrive\\Desktop\\Richard\\FPUNA\\Python\\Wordcloud\\feedback.xlsx", "Hoja1")
stw_spanish = obtener_stopwords()
text_tokens = array([])

#AGREGAR MAS PALABRAS AL STOPWORD
stw_spanish.extend(['ser', 'cierto'])

#PARTE PRINCIPAL
base = df[df['FEEDBACK'].notnull()]['FEEDBACK']

for a in base:
    text_tokens = append(text_tokens, word_tokenize(procesar_linea_texto(a)))

text_tokens = eliminar_palabras(text_tokens)

texto_nltk = Text(text_tokens)

frec_words = FreqDist(texto_nltk).most_common(20)

wc = DataFrame(frec_words, columns=['Word', 'Frecuency'])
wc.to_excel("data.xlsx")
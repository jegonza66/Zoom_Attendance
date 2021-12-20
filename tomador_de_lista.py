from PIL import Image
import pytesseract
import copy
import numpy as np
import pandas as pd
import pathlib
import os.path, time
import datetime as dt


def token(palabra, base): # Genera una base (diccionario) con las letras de esa palabra
    for i in range(len(palabra)):
        base[palabra[i]] = 0
    return base

def armar_base():
    base = {}
    abcstring = 'a'
    for i in range(32, 253, 1):
        abcstring += chr(i)
    abcstring += '‘' # Agrego la ‘
    abcstring += '’' # Agrego la ’
    abcstring += '“' # Agrego la “
    abcstring += '”' # Agrego la ”

    base = token(abcstring, base)
    return base

def basePalabra(palabra, base): # Cuenta cuantas letras de la base tiene la palabra
    basePal = copy.deepcopy(base)
    Error = 0
    for i in range(len(palabra)):
        try:
            basePal[palabra[i]] += 1
        except KeyError:
            Error = 1
            break
    if Error:
        return 0
    else:
        return basePal

def comparacion(palabra1, palabra2, base): # Devuelve distancia entre palabras:
    vector1 = basePalabra(palabra1, base)
    vector2 = basePalabra(palabra2, base)
    dist = 0
    if vector1 and vector2:
        for i in base.keys():
            dist += abs(vector1[i] - vector2[i])
    else:
        dist = 100
    return dist

def seleccion_palabra (palabra_1, lista, distancias, cota): # Selecciona la palabra mas parecida, sino la hay dice el error
	if min(distancias) <= cota:
		minimos = np.where(np.array(distancias) == min(distancias))[0]
		if len(minimos) > 1:
			x = str([lista[minimos[i]] for i in range(len(minimos))])
			print('\nEl nombre \033[1;31;49m {} \033[1;37;49m coincide con mas de una palabra ({})\n'.format(palabra_1, x))
		else: return lista[minimos[0]]
	else:
		print('\nEl nombre \033[1;31;49m {} \033[1;37;49m no tiene coincidencias\n'.format(palabra_1))

def correccion_palabras(ingresadas, lista, cota = 15): # Devuelve la lista de palabras corregidas
	base = armar_base()
	palabras_ingresadas = []
	for i in range(len(ingresadas)):
		distancias = []
		for j in range(len(lista)):
			distancias.append(comparacion(ingresadas[i], lista[j], base))
		supuesta_palabra = seleccion_palabra(ingresadas[i], lista, distancias, cota)
		if type(supuesta_palabra) == str:
			palabras_ingresadas.append(supuesta_palabra)
	return palabras_ingresadas

def limpiar_lista(lista, mas_corto):
    lista = [x.upper() for x in lista if len(x) >= mas_corto]
    return lista

# path = pathlib.Path(r'C:\Users\Joaco\Desktop\joaco\Colegios\Almafuerte\5B\Asistencia')
path = pathlib.Path(__file__).parent.absolute()
imagenes = list(path.glob('*.png'))+list(path.glob('*.jpg'))
tupla = []
for imagen in imagenes:
    fecha = dt.datetime.strptime(time.ctime(os.path.getctime(imagen)), '%c').date()
    tupla.append((imagen,fecha))
tupla = sorted(tupla, key=lambda x: x[1])
imagenes = [tupla[i][0] for i in range(len(tupla))]
archivo = list(path.glob('*.xlsx'))

#CARGO LISTA DE EXCEL
print('\nLeyendo Excel...\n')
df = pd.read_excel(archivo[0])
df = pd.DataFrame(df['ALUMNO'])
lista_alumnos = df.ALUMNO
lista_alumnos = limpiar_lista(lista_alumnos, 0)
for i in range(len(lista_alumnos)): df['ALUMNO'][i] = lista_alumnos[i]
largo_nombres = [len(lista_alumnos[i]) for i in range(len(lista_alumnos))]
tamaño_nombre_corto = min(largo_nombres)-4 # FILTRO POR NOMBRE MÁS CORTO

print('\nAnalizando imágenes...\n')
for imagen in imagenes:

    #ANALIZO SCREENSHOT
    fecha_str = time.ctime(os.path.getctime(imagen))
    fecha = dt.datetime.strptime(fecha_str, '%c').date()
    try:df.insert(loc=len(df.columns), column=str(fecha), value =0)
    except: pass
    value = Image.open(imagen) #ABRO IMAGEN
    texto_imagen = pytesseract.image_to_string(value,config= r'--tessdata-dir "C:\Program Files (x86)\Tesseract-OCR\tessdata"')
    alumnos_imagen = texto_imagen.splitlines() #SEPARO STR EN LISTA
    alumnos_imagen = limpiar_lista(alumnos_imagen, tamaño_nombre_corto) #LIMPIO LA LISTA DE NOMBRES

    alumnos_asistentes = correccion_palabras(alumnos_imagen, lista_alumnos)

    for alumno in alumnos_asistentes:
        fila = df.index[df['ALUMNO'] == alumno][0]
        df[str(fecha)][fila] = 1


df.insert(loc=len(df.columns), column='TOTAL', value =0)
df['TOTAL'] = df.sum(axis=1)

df.to_excel(archivo[0])
input("Press Enter to exit...")


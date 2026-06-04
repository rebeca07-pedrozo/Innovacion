# ---- DIAGNÓSTICO: ver cómo viene realmente el archivo ----
archivos = listar_txts(ID_CARPETA_DRIVE)
print("Archivos encontrados:", [a['name'] for a in archivos])

texto = descargar_texto(archivos[0]['id'])
lineas = texto.splitlines()

print("\n--- Primeras 5 líneas (repr, para ver tabs/espacios) ---")
for l in lineas[:5]:
    print(repr(l))

print("\n--- Conteo de posibles separadores en la 1ra línea ---")
h = lineas[0]
for nombre, ch in [("TAB", "\t"), ("punto y coma", ";"), ("coma", ","), ("pipe", "|")]:
    print(f"{nombre}: {h.count(ch)}")


# ---- DIAGNÓSTICO: ver la estructura REAL de las filas de datos ----
archivos = listar_txts(ID_CARPETA_DRIVE)
texto = descargar_texto(archivos[0]['id'])
lineas = texto.splitlines()

h = next(i for i, l in enumerate(lineas) if 'COD_AUX' in l)
print("== ENCABEZADO ==")
print(repr(lineas[h]))
print("\n== LÍNEAS SIGUIENTES (encabezado partido + guiones + primeros datos) ==")
for i in range(h, h + 9):
    if i < len(lineas):
        print(f"[{i}] len={len(lineas[i])}  tabs={lineas[i].count(chr(9))}")
        print("   ", repr(lineas[i]))


import os

carpeta_entrada = "/content/drive/My Drive/Optimizacion-Premios/2026/Entrada/"

print("¿Existe la carpeta?:", os.path.isdir(carpeta_entrada))
print("\nContenido REAL de la carpeta:")
for f in os.listdir(carpeta_entrada):
    print("  ->", repr(f))   # repr muestra espacios ocultos y la extensión real
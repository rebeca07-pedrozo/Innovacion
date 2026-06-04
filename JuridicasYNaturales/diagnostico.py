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
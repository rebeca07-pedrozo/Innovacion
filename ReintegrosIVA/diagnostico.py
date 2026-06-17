# ============================================================================
#  INSPECTOR DEL ARCHIVO PLANO  —  para entender la estructura y por qué falla
# ============================================================================
import re

# 👉 Si hace falta, ajusta estos dos:
ENCODING = 'utf-8'        # cambia a 'latin-1' si ves acentos/caracteres raros
N_LINEAS = 8              # cuántas líneas de muestra quieres ver

print("="*70)
print("PARTE 1 — ¿CÓMO LUCEN LAS PRIMERAS LÍNEAS? (caracteres exactos)")
print("="*70)
lineas_muestra = []
with open(RUTA_TXT, 'r', encoding=ENCODING, errors='replace') as f:
    for i, linea in enumerate(f):
        linea = linea.rstrip('\n').rstrip('\r')
        lineas_muestra.append(linea)
        print(f"\n[Línea {i+1}] longitud = {len(linea)} caracteres")
        print("   Contenido:", repr(linea[:200]))   # repr revela espacios, tabs, etc.
        if i + 1 >= N_LINEAS:
            break

print("\n" + "="*70)
print("PARTE 2 — ¿QUÉ SEPARADOR USA?")
print("="*70)
ejemplo = lineas_muestra[0] if lineas_muestra else ""
for nombre, car in [("Tabulación (\\t)", "\t"), ("Punto y coma (;)", ";"),
                    ("Coma (,)", ","), ("Barra (|)", "|"),
                    ("Espacios múltiples", "  ")]:
    cuenta = ejemplo.count(car)
    print(f"   {nombre:<22} aparece {cuenta} vez/veces en la 1ª línea")
print("   → Si todos dan 0 o casi, el archivo es de ANCHO FIJO (columnas pegadas).")

print("\n" + "="*70)
print("PARTE 3 — ¿CÓMO ESTÁN LAS CUENTAS DEL EXCEL?")
print("="*70)
try:
    print("   Cantidad de cuentas:", len(cuentas_set))
    muestra_cuentas = list(cuentas_set)[:8]
    for c in muestra_cuentas:
        print(f"      '{c}'  (longitud {len(c)})")
except NameError:
    print("   ⚠ No existe 'cuentas_set'. Corre primero el bloque que lee el Excel.")
    muestra_cuentas = []

print("\n" + "="*70)
print("PARTE 4 — PROBAR UNA CUENTA CONTRA EL ARCHIVO (3 métodos)")
print("="*70)
if muestra_cuentas:
    patron_num = re.compile(r'\d+')
    # Probamos cada cuenta de muestra hasta encontrar alguna que aparezca
    for cuenta in muestra_cuentas:
        hit_token = hit_sub = hit_strip = None
        with open(RUTA_TXT, 'r', encoding=ENCODING, errors='replace') as f:
            for linea in f:
                # método A: número exacto (tokens)
                if hit_token is None and cuenta in set(patron_num.findall(linea)):
                    hit_token = linea.rstrip('\n')
                # método B: "contiene" tal cual
                if hit_sub is None and cuenta in linea:
                    hit_sub = linea.rstrip('\n')
                # método C: "contiene" quitando ceros a la izquierda de la cuenta
                if hit_strip is None and cuenta.lstrip('0') and cuenta.lstrip('0') in linea:
                    hit_strip = linea.rstrip('\n')
                if hit_token and hit_sub and hit_strip:
                    break
        print(f"\n   Cuenta probada: '{cuenta}'")
        print("      A) número exacto (tokens)     :", "✅" if hit_token else "❌")
        print("      B) contiene (substring)        :", "✅" if hit_sub else "❌")
        print("      C) contiene sin ceros izq.     :", "✅" if hit_strip else "❌")
        if hit_sub:
            print("      Línea donde aparece (B):", repr(hit_sub[:200]))
            pos = hit_sub.find(cuenta)
            print(f"      Posición del texto en la línea: carácter {pos}")
        if hit_token or hit_sub or hit_strip:
            break   # ya encontramos una cuenta de ejemplo que sí está; suficiente
    else:
        print("\n   ⚠ Ninguna de las cuentas de muestra apareció con ningún método.")
        print("     Posibles causas: las cuentas no están en ESTE archivo, o el Excel")
        print("     está leyendo otra columna, o la codificación es distinta.")
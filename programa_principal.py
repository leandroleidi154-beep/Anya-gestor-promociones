from leer_promociones import leer_promociones
from leer_stock import leer_stock
from cruzar import cruzar_promociones_stock
from formato import generar_excel


archivo_promociones = "AGOSTO.xlsx"
archivo_stock = "Stock_tabla_prueba.xlsx"


print("Leyendo promociones...")

promociones = leer_promociones(
    archivo_promociones
)

print(f"Promociones encontradas: {len(promociones)}")


print("\nLeyendo Stock...")

stock = leer_stock(
    archivo_stock
)

print(f"Productos de Stock encontrados: {len(stock)}")


print("\nCruzando promociones con Stock...")

resultado = cruzar_promociones_stock(
    promociones,
    stock
)

print(f"Coincidencias encontradas: {len(resultado)}")


ARCHIVO_RESULTADO = "resultado.xlsx"

generar_excel(resultado, ARCHIVO_RESULTADO)

print("\nProceso terminado.")
print(f"Archivo creado: {ARCHIVO_RESULTADO}")

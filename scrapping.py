import requests
from bs4 import BeautifulSoup
import pandas as pd

def extraer_libros():
    url = "http://books.toscrape.com/"
    headers = {"User-Agent": "Mozilla/5.0"}
    
    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        print(f"Error al acceder a la página: {response.status_code}")
        return []

    soup = BeautifulSoup(response.text, "html.parser")
    libros = []

    # Extraer cada elemento de libro en la página
    for item in soup.find_all("article", class_="product_pod"):
        titulo = item.h3.a["title"]
        precio_texto = item.find("p", class_="price_color").text
        # Limpiar el precio convirtiéndolo a número flotante
        precio = float(precio_texto.replace("£", "").replace("Â", ""))
        
        disponibilidad = item.find("p", class_="instock availability").text.strip()
        
        libros.append({
            "Título": titulo,
            "Precio (£)": precio,
            "Disponibilidad": disponibilidad
        })

    return libros

def exportar_a_excel(datos, nombre_archivo="reporte_libros.xlsx"):
    df = pd.DataFrame(datos)
    
    # Filtrar solo productos con precio menor a £30 (Lógica de negocio aplicada)
    df_filtrado = df[df["Precio (£)"] < 30.0].sort_values(by="Precio (£)")

    # Exportar a Excel con formato
    with pd.ExcelWriter(nombre_archivo, engine="openpyxl") as writer:
        df_filtrado.to_excel(writer, index=False, sheet_name="Ofertas < £30")

    print(f"¡Éxito! Archivo generado: {nombre_archivo}")

if __name__ == "__main__":
    datos_libros = extraer_libros()
    if datos_libros:
        exportar_a_excel(datos_libros)
![Python](https://img.shields.io/badge/Python-3.8%2B-3776AB?style=flat&logo=python&logoColor=white) ![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey)

![GUI](https://img.shields.io/badge/GUI-Tkinter-red?style=flat)
![Status](https://img.shields.io/badge/Status-Development-orange?style=flat) 

# 📦 Csomagkövető Automatizáció (Kuehne+Nagel)

Ez a projekt automatizálja a csomagok (Kuehne+Nagel és GLS) nyomon követését. A szkript egy Excel fájlból beolvassa a csomagszámokat és az irányítószámokat, majd a háttérben (Web Scraping és belső API hívások segítségével) lekérdezi a csomagok legfrissebb útvonal-adatait és státuszát.

## Főbb funkciók

* **Tömeges feldolgozás:** Csomagadatok automatikus beolvasása `.xlsx` (Excel) fájlból.
* **Kétlépcsős adatnyerés:** Belső azonosítók (internal_id) kinyerése a weboldal HTML forráskódjából (Regex), majd a rejtett API végpontok meghívása a részletes adatokért.
* **Dátumok kinyerése:** A JSON válaszok feldolgozása a legfrissebb/legelső mérföldkövek (milestones) megtalálásához.

## Rendszerkövetelmények

A futtatáshoz **Python 3.x** környezet vagy a legfrissebb kiadású futtatható állomány szükséges. 

A külső könyvtárak telepítéséhez futtasd a parancssorban/terminálban a következő parancsot:

```bash
pip install requests openpyxl
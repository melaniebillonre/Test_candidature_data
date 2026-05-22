import pandas as pd

df = pd.read_csv("Base de données initiale.csv" , sep=";", decimal=",")
# Lire le fichier csv, avec séparateur ";" car format français

df = df.rename (columns={
    "Date de vente"  : "Date",
    "Produit"        : "Fruit",
    "Type"           : "Catégorie",
    "Quantité (kg)"  : "Quantité vendue (kg)",
    "Prix/kg (€)"    : "Prix unitaire (€ / kg)",
    "CA (€)"         : "Chiffre d'affaires (€)"
})
# Renommer les colonnes pour correspondre au excel demandé

df = df[["Date", "Fruit", "Catégorie", "Quantité vendue (kg)", "Prix unitaire (€ / kg)", "Chiffre d'affaires (€)"]]
# Remettre les colonne dans le même ordre que l'excel demandé

df.to_excel("Base de données à mettre à jour.xlsx", index=False, sheet_name="Ventes fruits" , engine="openpyxl")
# Remplace les anciennes données du fichier Excel par celles du CSV
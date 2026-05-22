import pandas as pd

df = pd.read_csv("Base de donnees initiale.csv" , sep=";", decimal=",")
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


# df.to_excel("Base de donnees a mettre a jour.xlsx", index=False, sheet_name="Ventes fruits" , engine="openpyxl")
# ce que j'avais mis initiallement mais remplacait completement le fichier excel et perdais la feuille Résumé


# A la place:
with pd.ExcelWriter("Base de donnees a mettre a jour.xlsx", engine="openpyxl", mode="w") as writer:
    df.to_excel(writer, sheet_name="Ventes fruits", index=False)
# permet de mettre à jour le fichier excel en gardant la feuille résumé,
# par contre on perd quand même la mise en page du tableau ventes fruits

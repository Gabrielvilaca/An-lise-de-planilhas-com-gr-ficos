import pandas as pd
import matplotlib.pyplot as plt
from matplotlib_venn import venn3

caminho_arquivo_firefox = "C:/Users/onde/esta/seu/arquivo/firefox.xlsx"
caminho_arquivo_zip = "C:/Users/onde/esta/seu/arquivo/zip.xlsx"
caminho_arquivo_office = "C:/Users/onde/esta/seu/arquivo/office.xlsx"

# Carregar as planilhas
planilha_firefox = pd.read_excel(caminho_arquivo_firefox)
planilha_7zip = pd.read_excel(caminho_arquivo_zip)
planilha_word = pd.read_excel(caminho_arquivo_office)

# Extrair os nomes dos computadores
computadores_firefox = set(planilha_firefox["Computador"])
computadores_7zip = set(planilha_7zip["Computador"])
computadores_word = set(planilha_word["Computador"])

# Calcular as interseções
todos_computadores = computadores_firefox | computadores_7zip | computadores_word
somente_firefox = computadores_firefox - computadores_7zip - computadores_word
somente_7zip = computadores_7zip - computadores_firefox - computadores_word
somente_word = computadores_word - computadores_7zip - computadores_firefox
firefox_7zip = computadores_firefox & computadores_7zip
firefox_word = computadores_firefox & computadores_word
zip_word = computadores_7zip & computadores_word
todos = computadores_firefox & computadores_7zip & computadores_word

# Gerar o diagrama de Venn
fig, ax = plt.subplots(figsize=(8, 8))

# Plotar o diagrama de Venn
venn3(subsets=(len(somente_firefox), len(somente_7zip), len(firefox_7zip),
               len(somente_word), len(firefox_word), len(zip_word),
               len(todos)),
      set_labels=('Firefox', '7-Zip', 'Word'))

# Ajustar o título e legendas
plt.title("Diagrama de Venn - Softwares Instalados")
plt.show()

# Imprimir máquinas em cada parte
print("Máquinas com Somente Firefox:")
print(somente_firefox)
print("\nMáquinas com Somente 7-Zip:")
print(somente_7zip)
print("\nMáquinas com Somente Word:")
print(somente_word)
print("\nMáquinas com Firefox & 7-Zip:")
print(firefox_7zip)
print("\nMáquinas com Firefox & Word:")
print(firefox_word)
print("\nMáquinas com 7-Zip & Word:")
print(zip_word)
print("\nMáquinas com Todos os 3:")
print(todos)

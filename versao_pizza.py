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

# Definir cores
cores = {"Firefox": "#003087", "7-Zip": "#FF7700", "Word": "#0070C0", "Todos": "#999999"}

# Desenhar áreas
ax.pie([len(todos), len(somente_firefox), len(somente_7zip), len(somente_word), 
        len(firefox_7zip), len(firefox_word), len(zip_word), len(todos)],
        labels=["Todos", "Somente Firefox", "Somente 7-Zip", "Somente Word", 
                "Firefox & 7-Zip", "Firefox & Word", "7-Zip & Word", "Todos os 3"],
        colors=[cores[k] for k in cores.keys()],
        startangle=90,
        autopct='%1.1f%%')

# Adicionar legendas
for label, color in cores.items():
    ax.annotate(label, xy=(0.8, 0.1 - 0.08 * list(cores.keys()).index(label)), 
                color=color, weight='bold')

# Ajustar o título e legendas
plt.title("Diagrama de Venn - Softwares Instalados")
plt.legend(loc="best")
plt.show()
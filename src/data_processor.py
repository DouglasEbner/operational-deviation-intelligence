import pandas as pd
import os
import re
import matplotlib.pyplot as plt

# ==================================================
# CONFIGURAÇÕES
# ==================================================

pasta = "data_saple"
status_validos = ["criterios"]

lista_status = []
lista_ofensores = []

# ==================================================
# LEITURA DOS ARQUIVOS
# ==================================================

for arquivo in sorted(os.listdir(pasta)):

    if not arquivo.endswith(".xlsb"):
        continue

    caminho_arquivo = os.path.join(pasta, arquivo)

    # Extrair DIA do nome do arquivo
    match = re.search(r'_(\d{1,2})_', arquivo)
    dia = int(match.group(1)) if match else None

    if dia is None:
        print(f"Dia não identificado no arquivo: {arquivo}")
        continue

    try:
        df = pd.read_excel(
            caminho_arquivo,
            sheet_name="sheet_name",
            engine="pyxlsb",
            header=2
        )
    except Exception as e:
        print(f"Erro ao ler {arquivo}: {e}")
        continue

    df.columns = df.columns.str.strip()

    if "Status Linhas" not in df.columns:
        print(f"Coluna 'Status Linhas' não encontrada em {arquivo}")
        continue

    # Filtrar status válidos
    df_filtrado = df[df["Status Linhas"].isin(status_validos)]

    # ==================================================
    # RESUMO STATUS POR UF
    # ==================================================

    resumo_temp = (
        df_filtrado
        .groupby(["UF", "Status Linhas"])
        .size()
        .reset_index(name="Quantidade")
    )
    resumo_temp["Dia"] = dia
    lista_status.append(resumo_temp)

    # ==================================================
    # OFENSORES (APENAS NOVOS)
    # ==================================================

    just_temp = (
        df_filtrado[df_filtrado["Status Linhas"] == "Nova"]
        .groupby(["UF", "Justificativas"])
        .size()
        .reset_index(name="Quantidade")
    )
    just_temp["Dia"] = dia
    lista_ofensores.append(just_temp)

# ==================================================
# CONSOLIDAÇÃO
# ==================================================

if not lista_status:
    raise ValueError("Nenhum dado válido foi encontrado.")

df_mes = pd.concat(lista_status, ignore_index=True)

df_ofensores = (
    pd.concat(lista_ofensores, ignore_index=True)
    if lista_ofensores else
    pd.DataFrame(columns=["UF", "Justificativas", "Quantidade", "Dia"])
)

# ==================================================
# EVOLUÇÃO DIÁRIA
# ==================================================

novos_mes = df_mes[df_mes["Status Linhas"] == "Nova"]

evolucao = (
    novos_mes
    .groupby("Dia")["Quantidade"]
    .sum()
    .reset_index()
    .sort_values("Dia")
)

evolucao_uf = (
    novos_mes
    .groupby(["Dia", "UF"])["Quantidade"]
    .sum()
    .reset_index()
    .sort_values(["UF", "Dia"])
)

# ==================================================
# GRÁFICO 1 - DIA ATUAL
# ==================================================

dia_atual = df_mes["Dia"].max()

df_hoje = df_mes[df_mes["Dia"] == dia_atual]
novos_hoje = (
    df_hoje[df_hoje["Status Linhas"] == "Nova"]
    .sort_values(by="Quantidade", ascending=False)
)

plt.figure()
plt.bar(novos_hoje["UF"], novos_hoje["Quantidade"])
plt.title(f"Desvios Novos - Dia {dia_atual}")
plt.xlabel("UF")
plt.ylabel("Quantidade")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig("grafico_dia_atual.png")
plt.close()

# ==================================================
# GRÁFICO 2 - EVOLUÇÃO POR UF
# ==================================================

plt.figure()

for uf in evolucao_uf["UF"].unique():
    dados_uf = evolucao_uf[evolucao_uf["UF"] == uf]
    plt.plot(dados_uf["Dia"], dados_uf["Quantidade"], marker="o", label=uf)

plt.title("Evolução Diária - Desvios Novos por UF")
plt.xlabel("Dia")
plt.ylabel("Total de Novos")
plt.xticks(sorted(evolucao["Dia"].unique()))
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.savefig("grafico_evolucao_uf.png")
plt.close

# ==================================================
# EXPORTAÇÃO EXCEL
# ==================================================

with pd.ExcelWriter("Consolidado_fevereiro_2026.xlsx") as writer:
    df_mes.to_excel(writer, sheet_name="Resumo_Status", index=False)
    df_ofensores.to_excel(writer, sheet_name="Ofensores_Dia", index=False)

# ==================================================
# DEBUG
# ==================================================

print("Dia Atual identificado:", dia_atual)
print(df_ofensores)

plt.show()

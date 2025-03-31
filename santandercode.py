# %% [markdown]
# Importação das bibliotecas necessárias:

# %%
import pandas as pd 
import os
import numpy as np
from rapidfuzz import process, fuzz
from IPython.display import display

# %%
# Configurar Pandas para exibir todas as linhas e colunas
pd.set_option("display.max_rows", None)  # Exibir todas as linhas
pd.set_option("display.max_columns", None)  # Exibir todas as colunas
pd.set_option("display.width", 1000)  # Ajustar a largura da saída

# %%
#Carregando as planilhas
caminho_erp =r"C:\Users\yan.fernandes\Downloads\Análise de Titulos de Cartão de Terceiros - SFR (1).csv"
caminho_santander =r"C:\Users\yan.fernandes\Downloads\Recebivel_Completos_9784485_20250101_20250131_17b71b847b834144add843c70f2feea1.xlsx"
# ✅ Função para carregar a planilha automaticamente (Excel ou CSV)
def carregar_planilha(caminho):
    if caminho.endswith(".csv"):
        return pd.read_csv(caminho, sep=";", encoding="latin1")  # Ajuste o separador se necessário
    else:
        return pd.read_excel(caminho, sheet_name="Detalhado") #Arquivo Santander tem uma aba que precisa ser considerada, poderia ser digitado manual!!!!
# ✅ Carregar as planilhas
df_erp = carregar_planilha(caminho_erp)
#print(df_erp.columns)
#print(df_erp.dtypes)
df_santander = carregar_planilha(caminho_santander)



# %%
#Limpando e organizando a planilha do Santander
#Função para limpar as primeiras linhas do Santander
def limpar_santander(df):
    df = df.iloc[6:].reset_index(drop=False)  # Remove as 7 primeiras linhas e reseta os índices
    df.columns = df.iloc[0]  # Define a primeira linha restante como cabeçalho
    df = df[1:].reset_index(drop=True)  # Remove a linha duplicada que virou cabeçalho
    return df

# ✅ Exibir informações básicas sobre os arquivos carregados
df_santander = limpar_santander(df_santander)
#display(df_santander)
df_santander = df_santander.filter(items=["EC CENTRALIZADOR", "DATA DE VENCIMENTO", "TIPO DE LANÇAMENTO", "PARCELAS", "AUTORIZAÇÃO", "NÚMERO COMPROVANTE DE VENDA (NSU)", "DATA DA VENDA","VALOR DA PARCELA", "VALOR LÍQUIDO"])

#Convertendo colunas para número
df_santander["VALOR LÍQUIDO"] = pd.to_numeric(df_santander["VALOR LÍQUIDO"], errors="coerce")
df_santander["VALOR DA PARCELA"] = pd.to_numeric(df_santander["VALOR DA PARCELA"], errors="coerce")


#Convertendo parcelas para números inteiros
df_santander[["PARCELA", "TOTAL_PARCELAS"]] = df_santander["PARCELAS"].str.extract(r"(\d+)\s+de\s+(\d+)") #Agora na planilha santander, o campo parcela vem em apenas 1 celula precisando separar em colunas.
df_santander["PARCELA"] = pd.to_numeric(df_santander["PARCELA"], errors="coerce")
df_santander["PARCELA"] = df_santander["PARCELA"].fillna(1).astype(int) #Essa linha converte o número da parcela do tipo float para interger porém quando a venda é no débito o mesmo vem zerado. Sendo assim optou-se por preencher esse campo como valor 1, o mesmo ocorre para quantidade de parcelas
df_santander["TOTAL_PARCELAS"] = pd.to_numeric(df_santander["TOTAL_PARCELAS"], errors="coerce")
df_santander["TOTAL_PARCELAS"] = df_santander["TOTAL_PARCELAS"].fillna(1).astype(int)

#Convertendo Data do pagamento e Data do lançamento para data
df_santander["DATA DA VENDA"] = pd.to_datetime(df_santander["DATA DA VENDA"], format="%d/%m/%Y", errors="coerce")
df_santander["DATA DE VENCIMENTO"] = pd.to_datetime(df_santander["DATA DE VENCIMENTO"], format="%d/%m/%Y", errors="coerce")

#Separando os valores de aluguel de máquina e cancelamento dos valores da GETNET.
df_cancelamento_venda = df_santander[df_santander["TIPO DE LANÇAMENTO"] == "Cancelamento/Chargeback"]
df_aluguel_maquina = df_santander[df_santander["TIPO DE LANÇAMENTO"] == "Aluguel/Tarifa"]

#Atualizando a tabela df_santander para todos os valores sem o aluguel de máquina, sem cancelamento e sem valores em branco
df_santander = df_santander[df_santander["TIPO DE LANÇAMENTO"] != "Cancelamento/Chargeback"]
df_santander = df_santander[df_santander["TIPO DE LANÇAMENTO"] != "Aluguel/Tarifa"]
df_santander = df_santander[df_santander["TIPO DE LANÇAMENTO"] != "Pagamento Realizado"]
df_santander = df_santander[df_santander["TIPO DE LANÇAMENTO"] != "Saldo Anterior"]
df_santander = df_santander[df_santander["TIPO DE LANÇAMENTO"].notna()]

display(df_santander.dtypes)
#display(df_santander)
#Totalizadores
valor_total_bruto = df_santander["VALOR DA PARCELA"].sum()
valor_total_liquido =  df_santander["VALOR LÍQUIDO"].sum()
valor_aluguel_maquina = df_aluguel_maquina["VALOR LÍQUIDO"].sum()
valor_cancelamento_venda = df_cancelamento_venda["VALOR LÍQUIDO"].sum()
valor_recebido_conta = valor_total_liquido - abs(valor_aluguel_maquina) - abs(valor_cancelamento_venda)

#df_cancelamento_venda.to_excel(r"C:\Users\regiel\Documents\PythonAppsSantander\CancelamentoVenda.xlsx", index=False)
#df_santander.to_excel(r"C:\Users\regiel\Documents\PythonAppsSantander\Santander.xlsx", index=False)

print(valor_total_bruto)
print(valor_total_liquido)
print(valor_aluguel_maquina)
print(valor_cancelamento_venda)
print(valor_recebido_conta)

# %%
#Selecionando as colunas desejadas
df_erp = df_erp.filter(items=["1o. Agrupamento", "Chave", "Numero", "NSU", "Autorização", "Emissão", "Correção", "Valor", "Vr Corrigido"])
#Convertendo colunas para os tipos corretos
#Convertendo colunas para número
df_erp["Valor"] = df_erp["Valor"].str.replace(",", ".", regex=True)
df_erp["Valor"] = pd.to_numeric(df_erp["Valor"], errors="coerce")
df_erp["Vr Corrigido"] = df_erp["Vr Corrigido"].str.replace(",", ".", regex=True)
df_erp["Vr Corrigido"] = pd.to_numeric(df_erp["Vr Corrigido"], errors="coerce")
#Convertendo colunas para data
df_erp["Emissão"] = pd.to_datetime(df_erp["Emissão"], format="%d/%m/%Y", errors="coerce")
df_erp["Correção"] = pd.to_datetime(df_erp["Correção"], format="%d/%m/%Y", errors="coerce")
#Transformando o campo Numero em Parcela e Total de Parcelas
# Criar as novas colunas extraindo os valores corretos da coluna "Numero"
df_erp["chcriacao"] = df_erp["Numero"].str.split("-").str[0]  # Antes do "-"
df_erp["Parcela"] = df_erp["Numero"].str.split("-").str[1].str.split("/").str[0]  # Entre "-" e "/"
df_erp["Total_Parcelas"] = df_erp["Numero"].str.split("/").str[1]  # Após "/"

# Converter as colunas de parcela para inteiro
df_erp["Parcela"] = pd.to_numeric(df_erp["Parcela"], errors="coerce").fillna(1).astype(int)
df_erp["Total_Parcelas"] = pd.to_numeric(df_erp["Total_Parcelas"], errors="coerce").fillna(1).astype(int)
df_erp = df_erp.filter(items=["1o. Agrupamento", "Chave", "chcriacao", "Parcela", "Total_Parcelas", "NSU", "Autorização", "Emissão", "Correção", "Valor", "Vr Corrigido"])

#Selecionando apenas os títulos da industria
df_erp_loja = df_erp[~df_erp["1o. Agrupamento"].isin(["LE SFR Indústria Ltda", "LE Protendidos"])].copy()


display(df_erp_loja.dtypes)
#display(df_erp_loja)


# %%
#Funções Conciliar por valor e data e conciliando buscando autorizações parecidas

def conciliar_por_data_e_valores(row, df_erp_base):
    #print("\n🔍 Buscando correspondência por data e valores para:", row["AUTORIZAÇÃO"])

    # 1️⃣ Filtra por datas com até 5 dias de diferença
    data_diferenca = (df_erp_base["Emissão"] - row["DATA DA VENDA"]).abs().dt.days
    #print(f"→ Diferença de dias entre 'Emissão' e 'DATA DA VENDA':\n{data_diferenca.describe()}")

    candidatos = df_erp_base[data_diferenca <= 1]
    #print(f"→ Candidatos com diferença de até 5 dias: {len(candidatos)}")

    # 2️⃣ Filtra por valor, parcela e total de parcelas
    candidatos = candidatos[
        ((candidatos["Valor"] - row["VALOR DA PARCELA"]).abs() <= 0.12
        ) &
        (candidatos["Parcela"] == row["PARCELA"]) &
        (candidatos["Total_Parcelas"] == row["TOTAL_PARCELAS"])
    ]

    #print(f"→ Candidatos com valores compatíveis e parcelas iguais: {len(candidatos)}")
    if not candidatos.empty:
        linha = candidatos.iloc[0]

        #print(f"✅ Conciliado com:\n"
        #      f"   Autorização: {linha['Autorização']}\n"
        #      f"   Valor ERP: {linha['Valor']} | Valor Banco: {row['VALOR DA PARCELA']}\n"
        #      f"   Data ERP: {linha['Emissão']} | Data Venda: {row['DATA DA VENDA']}\n"
        #      f"   Parcela: {linha['Parcela']} | Total Parcelas: {linha['Total_Parcelas']}\n")

        return pd.Series([
            linha["Autorização"],
            linha["Chave"],
            linha["Valor"],
            "Conciliado por Data e Valores",
            10
        ])
    
    #print("❌ Nenhuma correspondência encontrada com os critérios de data, valor e parcelas.")
    return pd.Series([None, None, None, "Não Conciliado", 99])

def encontrar_melhor_correspondencia_com_pontuacao(row, df_origem, coluna_erp):
    correspondencias = process.extract(
        str(row["AUTORIZAÇÃO"]),
        df_origem[coluna_erp].astype(str),
        scorer=fuzz.ratio,
        limit=10
    )

    correspondencias_validas = [(texto, score, idx) for texto, score, idx in correspondencias if score >= 80]

    #print(f"\n🔍 Buscando correspondência para: {row['AUTORIZAÇÃO']}")
    #print("Correspondências válidas (score >= 80):", correspondencias_validas)

    if not correspondencias_validas:
        return pd.Series([None, None, None, "Não Conciliado", 99])

    melhor_resultado = None
    menor_pontuacao = float("inf")

    for melhor_correspondencia, melhor_pontuacao, _ in correspondencias_validas:
        filtro = df_origem[df_origem[coluna_erp] == melhor_correspondencia]

        if filtro.empty:
            #print(f"⚠ Correspondência '{melhor_correspondencia}' não encontrada no DataFrame.")
            continue

        # 🔄 Itera sobre todas as linhas com o mesmo valor
        for _, linha_correspondente in filtro.iterrows():
            valor_erp = linha_correspondente["Valor"]
            data_erp = linha_correspondente["Emissão"]
            parcela_erp = linha_correspondente["Parcela"]
            total_parcelas_erp = linha_correspondente["Total_Parcelas"]

            status = ["Conciliado"]
            pontuacao = 0

            if abs(row["VALOR DA PARCELA"] - valor_erp) > 0.10:
                status.append("Divergência de Valor")
                pontuacao += 15

            if abs((row["DATA DA VENDA"] - data_erp).days) > 1:
                status.append("Divergência de Data")
                pontuacao += 5

            if row["PARCELA"] != parcela_erp:
                status.append("Divergência de Parcela")
                pontuacao += 10

            if row["TOTAL_PARCELAS"] != total_parcelas_erp:
                status.append("Divergência de Total de Parcelas")
                pontuacao += 15

            #print("🔸 Analisando:", melhor_correspondencia)
            #print("    → Valor ERP:", valor_erp)
            #print("    → Data ERP:", data_erp)
            #print("    → Parcela ERP:", parcela_erp)
            #print("    → Total Parcelas ERP:", total_parcelas_erp)
            #print("    → Status:", status)
            #print("    → Pontuação calculada:", pontuacao)
            #print("    → Menor pontuação até agora:", menor_pontuacao)

            if pontuacao < menor_pontuacao:
                menor_pontuacao = pontuacao
                melhor_resultado = (
                    linha_correspondente[coluna_erp],
                    linha_correspondente["Chave"],
                    valor_erp,
                    " e ".join(status) if len(status) > 1 else status[0],
                    pontuacao
                )

    if melhor_resultado:
        #print("✅ Melhor resultado escolhido:", melhor_resultado)
        return pd.Series(melhor_resultado)
    else:
        #print("❌ Nenhuma correspondência com pontuação aceitável.")
        return pd.Series([None, None, None, "Não Conciliado", 99])
    
def encontrar_melhor_correspondencia_com_pontuacao_nsu(row, df_origem):
    correspondencias = process.extract(
        str(row["NÚMERO COMPROVANTE DE VENDA (NSU)"]),
        df_origem["NSU"].astype(str),
        scorer=fuzz.ratio,
        limit=10
    )

    correspondencias_validas = [(texto, score, idx) for texto, score, idx in correspondencias if score >= 80]

    print(f"\n🔍 Buscando correspondência para: {row['NÚMERO COMPROVANTE DE VENDA (NSU)']}")
    print("Correspondências válidas (score >= 80):", correspondencias_validas)

    if not correspondencias_validas:
        return pd.Series([None, None, None, "Não Conciliado", 99])

    melhor_resultado = None
    menor_pontuacao = float("inf")

    for melhor_correspondencia, melhor_pontuacao, _ in correspondencias_validas:
        filtro = df_origem[df_origem["NSU"] == melhor_correspondencia]

        if filtro.empty:
            print(f"⚠ Correspondência '{melhor_correspondencia}' não encontrada no DataFrame.")
            continue

        # 🔄 Itera sobre todas as linhas com o mesmo valor
        for _, linha_correspondente in filtro.iterrows():
            valor_erp = linha_correspondente["Valor"]
            data_erp = linha_correspondente["Emissão"]
            parcela_erp = linha_correspondente["Parcela"]
            total_parcelas_erp = linha_correspondente["Total_Parcelas"]

            status = ["Conciliado"]
            pontuacao = 0

            if abs(row["VALOR DA PARCELA"] - valor_erp) > 0.09:
                status.append("Divergência de Valor")
                pontuacao += 15

            if abs((row["DATA DA VENDA"] - data_erp).days) > 1:
                status.append("Divergência de Data")
                pontuacao += 5

            if row["PARCELA"] != parcela_erp:
                status.append("Divergência de Parcela")
                pontuacao += 10

            if row["TOTAL_PARCELAS"] != total_parcelas_erp:
                status.append("Divergência de Total de Parcelas")
                pontuacao += 15

            print("🔸 Analisando:", melhor_correspondencia)
            print("    → Valor ERP:", valor_erp)
            print("    → Data ERP:", data_erp)
            print("    → Parcela ERP:", parcela_erp)
            print("    → Total Parcelas ERP:", total_parcelas_erp)
            print("    → Status:", status)
            print("    → Pontuação calculada:", pontuacao)
            print("    → Menor pontuação até agora:", menor_pontuacao)

            if pontuacao < menor_pontuacao:
                menor_pontuacao = pontuacao
                melhor_resultado = (
                    linha_correspondente["NSU"],
                    linha_correspondente["Chave"],
                    valor_erp,
                    " e ".join(status) if len(status) > 1 else status[0],
                    pontuacao
                )

    if melhor_resultado:
        print("✅ Melhor resultado escolhido:", melhor_resultado)
        return pd.Series(melhor_resultado)
    else:
        print("❌ Nenhuma correspondência com pontuação aceitável.")
        return pd.Series([None, None, None, "Não Conciliado", 99])

# %%
#Remover da Planilha Santander os Títulos que foram cancelados
# 1️⃣ Criar coluna auxiliar com valor absoluto da parcela
df_santander["VALOR_ABS"] = df_santander["VALOR DA PARCELA"].abs()
df_cancelamento_venda["VALOR_ABS"] = df_cancelamento_venda["VALOR DA PARCELA"].abs()

# 2️⃣ Criar chave composta: AUTORIZAÇÃO + VALOR_ABS
df_santander["CHAVE_CONCILIACAO"] = df_santander["AUTORIZAÇÃO"].astype(str) + "_" + df_santander["VALOR_ABS"].astype(str)
df_cancelamento_venda["CHAVE_CONCILIACAO"] = df_cancelamento_venda["AUTORIZAÇÃO"].astype(str) + "_" + df_cancelamento_venda["VALOR_ABS"].astype(str)

# 3️⃣ Verificar chaves em comum
chaves_comuns = set(df_santander["CHAVE_CONCILIACAO"]) & set(df_cancelamento_venda["CHAVE_CONCILIACAO"])
print("🔍 Chaves encontradas em comum:", len(chaves_comuns))
print("👀 Exemplo de chaves comuns:", list(chaves_comuns)[:5])

# 4️⃣ Filtrar as linhas da df_santander que estão na lista de cancelamentos
filtro_cancelados = df_santander["CHAVE_CONCILIACAO"].isin(df_cancelamento_venda["CHAVE_CONCILIACAO"])
print("✂️ Linhas encontradas para recorte:", filtro_cancelados.sum())

# 5️⃣ Copiar essas linhas
df_cancelados_encontrados = df_santander[filtro_cancelados].copy()
#df_cancelados_encontrados["Status"] = "Cancelado"
#df_cancelados_encontrados["Pontuação"] = 101

# 6️⃣ Adicionar ao df_cancelamento_venda
df_cancelamento_venda = pd.concat([df_cancelamento_venda, df_cancelados_encontrados], ignore_index=True)

# 7️⃣ Remover da df_santander
df_santander = df_santander[~filtro_cancelados].copy()

# 8️⃣ Resultado final
print("✅ Linhas restantes em df_santander:", len(df_santander))
print("✅ Linhas totais em df_cancelamento_venda:", len(df_cancelamento_venda))
#display(df_cancelamento_venda)
#df_cancelamento_venda.to_excel(r"C:\Users\regiel\Documents\PythonAppsSantander\CancelamentoVenda02.xlsx", index=False)


# %%
#Primeira conciliação, procurar com todos os campos iguais. Será se vale a pena separar por loja?

df_primeira_conciliacao = df_santander.merge(
    df_erp_loja[["Chave", "Autorização", "NSU", "Emissão", "Parcela", "Total_Parcelas", "Valor"]],
    left_on=["AUTORIZAÇÃO", "DATA DA VENDA", "PARCELA", "TOTAL_PARCELAS"],
    right_on=["Autorização", "Emissão", "Parcela", "Total_Parcelas"],
    how="left"
)

#display(df_primeira_conciliacao) #Apenas debug 

#Criar totalizadores - Acompanhar quanto falta a conciliar do arquivo original

df_pri_conc_nao_conc = df_primeira_conciliacao[df_primeira_conciliacao["Chave"].isna()]
df_primeira_conciliacao = df_primeira_conciliacao[df_primeira_conciliacao["Chave"].notna()]
df_primeira_conciliacao["Status"] = "Conciliado"
df_primeira_conciliacao["Pontuação"] = 0
display("Aqui está a primeira conciliação:")
display(df_primeira_conciliacao)


# %%
df_segunda_conciliacao = df_pri_conc_nao_conc.filter(items=["EC CENTRALIZADOR", "DATA DE VENCIMENTO", "TIPO DE LANÇAMENTO", "PARCELAS", "AUTORIZAÇÃO", "NÚMERO COMPROVANTE DE VENDA (NSU)", "DATA DA VENDA","VALOR DA PARCELA", "VALOR LÍQUIDO", "PARCELA", "TOTAL_PARCELAS"])
df_segunda_conciliacao[["Melhor Autorização", "Chave ERP", "Valor ERP", "Status", "Pontuação"]] = df_segunda_conciliacao.apply(
    lambda row: conciliar_por_data_e_valores(row, df_erp),
    axis=1
)

# %%
df_terceira_conciliacao = df_segunda_conciliacao[df_segunda_conciliacao["Chave ERP"].isna()]
df_segunda_conciliacao = df_segunda_conciliacao[df_segunda_conciliacao["Chave ERP"].notna()]
#display(df_terceira_conciliacao)
#display(df_segunda_conciliacao)

# %%
df_terceira_conciliacao[["Melhor Autorização", "Chave ERP", "Valor ERP", "Status", "Pontuação"]] = df_terceira_conciliacao.apply(
    lambda row: encontrar_melhor_correspondencia_com_pontuacao(row, df_erp, "Autorização"),
    axis=1
)
df_quarta_conciliacao = df_terceira_conciliacao[df_terceira_conciliacao["Pontuação"] >= 6]
#display(df_quarta_conciliacao)
df_terceira_conciliacao = df_terceira_conciliacao[df_terceira_conciliacao["Pontuação"] < 6]
#display(df_terceira_conciliacao)

# %%
#df_terceira_conciliacao.to_excel(r"C:\Users\regiel\Documents\PythonAppsSantander\ConciliacaoTerceira.xlsx", index=False)
#df_erp_loja.to_excel(r"C:\Users\regiel\Documents\PythonAppsSantander\ConciliacaoERPLoja.xlsx", index=False)
display(df_terceira_conciliacao)
display(df_quarta_conciliacao)

# %%
#df_quarta_conciliacao.to_excel(r"C:\Users\yan.fernandes\Downloads\QuartaConciliacao.xlsx", index=False)
df_segunda_conciliacao.to_excel(r"C:\Users\yan.fernandes\Downloads\SegundaConciliacao.xlsx", index=False)




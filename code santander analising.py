# Importação das bibliotecas necessárias:
# %%
import pandas as pd 
import os
import numpy as np
from rapidfuzz import process, fuzz
from IPython.display import display
from openpyxl.styles import Font, Alignment

# %%
# Configurar Pandas para exibir todas as linhas e colunas
pd.set_option("display.max_rows", None)  # Exibir todas as linhas
pd.set_option("display.max_columns", None)  # Exibir todas as colunas
pd.set_option("display.width", 1000)  # Ajustar a largura da saída

# %%
#Carregando as planilhas
caminho_erp =r"C:\Users\yan.fernandes\Documents\conciliação santander\Análise de Titulos de Cartão de Terceiros - SFR.csv"
caminho_santander =r"C:\Users\yan.fernandes\Documents\conciliação santander\Recebivel_Completos_9784485_20250101_20250131_17b71b847b834144add843c70f2feea1.xlsx"
#  Função para carregar a planilha automaticamente (Excel ou CSV)
def carregar_planilha(caminho):
    if caminho.endswith(".csv"):
        return pd.read_csv(caminho, sep=";", encoding="latin1", dtype={"NSU": str})  # Ajuste o separador se necessário
    else:
        return pd.read_excel(caminho, sheet_name="Detalhado", dtype={"NÚMERO COMPROVANTE DE VENDA (NSU)": str}) #Arquivo Santander tem uma aba que precisa ser considerada, poderia ser digitado manual!!!!
#  Carregar as planilhas
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

#  Exibir informações básicas sobre os arquivos carregados
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
quantidade_titulos_santander = df_santander["VALOR DA PARCELA"].count()
valor_total_liquido =  df_santander["VALOR LÍQUIDO"].sum()
valor_aluguel_maquina = df_aluguel_maquina["VALOR LÍQUIDO"].sum()
valor_cancelamento_venda = df_cancelamento_venda["VALOR LÍQUIDO"].sum()
quantidade_titulos_cancelados = df_cancelamento_venda["VALOR LÍQUIDO"].count()
valor_recebido_conta = valor_total_liquido - abs(valor_aluguel_maquina) - abs(valor_cancelamento_venda)

#df_cancelamento_venda.to_excel(r"C:\Users\regiel\Documents\PythonAppsSantander\CancelamentoVenda.xlsx", index=False)
#df_santander.to_excel(r"C:\Users\regiel\Documents\PythonAppsSantander\Santander.xlsx", index=False)

print(valor_total_bruto)
print(quantidade_titulos_santander)
print(valor_total_liquido)
print(valor_aluguel_maquina)
print(valor_cancelamento_venda)
print(quantidade_titulos_cancelados)
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
    #print("\n Buscando correspondência por data e valores para:", row["AUTORIZAÇÃO"])

    # 1️ Filtra por datas com até 5 dias de diferença
    data_diferenca = (df_erp_base["Emissão"] - row["DATA DA VENDA"]).abs().dt.days
    #print(f"→ Diferença de dias entre 'Emissão' e 'DATA DA VENDA':\n{data_diferenca.describe()}")

    candidatos = df_erp_base[data_diferenca <= 5]
    #print(f"→ Candidatos com diferença de até 5 dias: {len(candidatos)}")

    # 2️ Filtra por valor, parcela e total de parcelas
    candidatos = candidatos[
        ((candidatos["Valor"] - row["VALOR DA PARCELA"]).abs() <= 0.20) &
        (candidatos["Parcela"] == row["PARCELA"]) &
        (candidatos["Total_Parcelas"] == row["TOTAL_PARCELAS"])
    ]

    #print(f"→ Candidatos com valores compatíveis e parcelas iguais: {len(candidatos)}")
    if not candidatos.empty:
        linha = candidatos.iloc[0]

        #print(f" Conciliado com:\n"
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
    
    #print(" Nenhuma correspondência encontrada com os critérios de data, valor e parcelas.")
    return pd.Series([None, None, None, "Não Conciliado", 99])

def encontrar_melhor_correspondencia_com_pontuacao(row, df_origem, coluna_erp):
    correspondencias = process.extract(
        str(row["AUTORIZAÇÃO"]),
        df_origem[coluna_erp].astype(str),
        scorer=fuzz.ratio,
        limit=10
    )

    correspondencias_validas = [(texto, score, idx) for texto, score, idx in correspondencias if score >= 80]

    #print(f"\n Buscando correspondência para: {row['AUTORIZAÇÃO']}")
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

        #  Itera sobre todas as linhas com o mesmo valor
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

            #print(" Analisando:", melhor_correspondencia)
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
        #print(" Melhor resultado escolhido:", melhor_resultado)
        return pd.Series(melhor_resultado)
    else:
        #print(" Nenhuma correspondência com pontuação aceitável.")
        return pd.Series([None, None, None, "Não Conciliado", 99])
    
def encontrar_melhor_correspondencia_com_pontuacao_nsu(row, df_origem):
    correspondencias = process.extract(
        str(row["NÚMERO COMPROVANTE DE VENDA (NSU)"]),
        df_origem["NSU"].astype(str),
        scorer=fuzz.ratio,
        limit=10
    )

    correspondencias_validas = [(texto, score, idx) for texto, score, idx in correspondencias if score >= 80]

    print(f"\n Buscando correspondência para: {row['NÚMERO COMPROVANTE DE VENDA (NSU)']}")
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

        #  Itera sobre todas as linhas com o mesmo valor
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

            print(" Analisando:", melhor_correspondencia)
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
        print(" Melhor resultado escolhido:", melhor_resultado)
        return pd.Series(melhor_resultado)
    else:
        print(" Nenhuma correspondência com pontuação aceitável.")
        return pd.Series([None, None, None, "Não Conciliado", 99])

def selecionar_melhor_por_pontuacao_com_autorizacao_e_nsu(row, df_erp_base, tolerancia_dias=5, tolerancia_valor=0.20, incluir_detalhes=False):
    candidatos = df_erp_base[
        (df_erp_base["Emissão"] - row["DATA DA VENDA"]).abs().dt.days <= tolerancia_dias
    ]

    candidatos = candidatos[
        (candidatos["Valor"] - row["VALOR DA PARCELA"]).abs() <= tolerancia_valor
    ]

    candidatos = candidatos[
        (candidatos["Parcela"] == row["PARCELA"]) &
        (candidatos["Total_Parcelas"] == row["TOTAL_PARCELAS"])
    ]

    if candidatos.empty:
        if incluir_detalhes:
            return pd.Series([None, None, None, None, None, None, "Não Conciliado", 999])
        else:
            return pd.Series([None, None, None, None, "Não Conciliado", 999])

    melhor_resultado = None
    menor_pontuacao = float("inf")

    for _, linha in candidatos.iterrows():
        dias_dif = abs((linha["Emissão"] - row["DATA DA VENDA"]).days)
        valor_dif = abs(linha["Valor"] - row["VALOR DA PARCELA"])

        aut_sant = str(row["AUTORIZAÇÃO"]).strip()
        aut_erp = str(linha["Autorização"]).strip()
        nsu_sant = str(row["NÚMERO COMPROVANTE DE VENDA (NSU)"]).strip()
        nsu_erp = str(linha["NSU"]).strip()

        if aut_sant == aut_erp or nsu_sant == nsu_erp:
            sim_autorizacao = 100
            sim_nsu = 100
        else:
            sim_autorizacao = fuzz.ratio(aut_sant, aut_erp)
            sim_nsu = fuzz.ratio(nsu_sant, nsu_erp)

        pontuacao = dias_dif * 100 + valor_dif * 100 + (200 - (sim_autorizacao + sim_nsu))

        if pontuacao < menor_pontuacao:
            menor_pontuacao = pontuacao
            melhor_resultado = (
                linha["Autorização"],
                linha["NSU"],
                linha["Chave"],
                linha["Valor"],
                dias_dif,
                valor_dif,
                "Conciliado por Similaridade",
                round(pontuacao, 2)
            )

    if melhor_resultado:
        if incluir_detalhes:
            return pd.Series(melhor_resultado)
        else:
            return pd.Series(melhor_resultado[:4] + melhor_resultado[-2:])  # sem dias/valor
    else:
        if incluir_detalhes:
            return pd.Series([None, None, None, None, None, None, "Não Conciliado", 999])
        else:
            return pd.Series([None, None, None, None, "Não Conciliado", 999])
        
def marcar_duplicados_com_pior_score(df, chave_col="Chave ERP", status_col="Status", pontuacao_col="Pontuação"):
    # 1️ Filtra linhas com chaves duplicadas
    duplicadas = df[df.duplicated(subset=[chave_col], keep=False)].copy()

    if duplicadas.empty:
        print(" Nenhuma chave duplicada encontrada.")
        return df

    print(f" Chaves duplicadas encontradas: {duplicadas[chave_col].nunique()}")

    # 2️ Ordena pela pontuação crescente (menor pontuação é a melhor)
    duplicadas_sorted = duplicadas.sort_values(pontuacao_col, ascending=True)

    # 3️ Marca como duplicado todas as duplicatas exceto a com menor pontuação
    duplicadas_marcadas = duplicadas_sorted.duplicated(subset=[chave_col], keep="first")

    # 4️ Atualiza status e pontuação das duplicadas com pior score
    df.loc[duplicadas_sorted[duplicadas_marcadas].index, status_col] = "Valor Duplicado Menor Score"
    df.loc[duplicadas_sorted[duplicadas_marcadas].index, pontuacao_col] = 998

    print(f" Linhas marcadas com 'Valor Duplicado Menor Score': {duplicadas_marcadas.sum()}")

    return df


# %%
#Remover da Planilha Santander os Títulos que foram cancelados
# 1️ Criar coluna auxiliar com valor absoluto da parcela
df_santander["VALOR_ABS"] = df_santander["VALOR DA PARCELA"].abs()
df_cancelamento_venda["VALOR_ABS"] = df_cancelamento_venda["VALOR DA PARCELA"].abs()

# 2️ Criar chave composta: AUTORIZAÇÃO + VALOR_ABS
df_santander["CHAVE_CONCILIACAO"] = df_santander["AUTORIZAÇÃO"].astype(str) + "_" + df_santander["VALOR_ABS"].astype(str)
df_cancelamento_venda["CHAVE_CONCILIACAO"] = df_cancelamento_venda["AUTORIZAÇÃO"].astype(str) + "_" + df_cancelamento_venda["VALOR_ABS"].astype(str)

# 3️ Verificar chaves em comum
chaves_comuns = set(df_santander["CHAVE_CONCILIACAO"]) & set(df_cancelamento_venda["CHAVE_CONCILIACAO"])
print(" Chaves encontradas em comum:", len(chaves_comuns))
print(" Exemplo de chaves comuns:", list(chaves_comuns)[:5])

# 4️ Filtrar as linhas da df_santander que estão na lista de cancelamentos
filtro_cancelados = df_santander["CHAVE_CONCILIACAO"].isin(df_cancelamento_venda["CHAVE_CONCILIACAO"])
print(" Linhas encontradas para recorte:", filtro_cancelados.sum())

# 5️ Copiar essas linhas
df_cancelados_encontrados = df_santander[filtro_cancelados].copy()
#df_cancelados_encontrados["Status"] = "Cancelado"
#df_cancelados_encontrados["Pontuação"] = 101

# 6️ Adicionar ao df_cancelamento_venda
df_cancelamento_venda = pd.concat([df_cancelamento_venda, df_cancelados_encontrados], ignore_index=True)

# 7️ Remover da df_santander
df_santander = df_santander[~filtro_cancelados].copy()

# 8️ Resultado final
print(" Linhas restantes em df_santander:", len(df_santander))
print(" Linhas totais em df_cancelamento_venda:", len(df_cancelamento_venda))
#display(df_cancelamento_venda)
#df_cancelamento_venda.to_excel(r"C:\Users\regiel\Documents\PythonAppsSantander\CancelamentoVenda02.xlsx", index=False)


# %%
df_primeira_conciliacao = df_santander
#df_segunda_conciliacao = df_pri_conc_nao_conc.filter(items=["EC CENTRALIZADOR", "DATA DE VENCIMENTO", "TIPO DE LANÇAMENTO", "PARCELAS", "AUTORIZAÇÃO", "NÚMERO COMPROVANTE DE VENDA (NSU)", "DATA DA VENDA","VALOR DA PARCELA", "VALOR LÍQUIDO", "PARCELA", "TOTAL_PARCELAS"])
df_segunda_conciliacao = df_primeira_conciliacao.filter(items=["EC CENTRALIZADOR", "DATA DE VENCIMENTO", "TIPO DE LANÇAMENTO", "PARCELAS", "AUTORIZAÇÃO", "NÚMERO COMPROVANTE DE VENDA (NSU)", "DATA DA VENDA","VALOR DA PARCELA", "VALOR LÍQUIDO", "PARCELA", "TOTAL_PARCELAS"])
df_segunda_conciliacao[["Autorização ERP", "NSU ERP", "Chave ERP", "Valor ERP", "Status", "Pontuação"]] = df_segunda_conciliacao.apply(
    lambda row: selecionar_melhor_por_pontuacao_com_autorizacao_e_nsu(row, df_erp),
    axis=1
)
#df_duplicados = df_segunda_conciliacao[df_segunda_conciliacao.duplicated(subset=["Chave ERP"], keep=False)].copy()
#df_segunda_conciliacao = df_segunda_conciliacao[~df_segunda_conciliacao.duplicated(subset=["Chave ERP"], keep=False)].copy()

# %%

df_terceira_conciliacao = df_segunda_conciliacao[df_segunda_conciliacao["Pontuação"] == 999].copy()
df_segunda_conciliacao = df_segunda_conciliacao[df_segunda_conciliacao["Pontuação"] != 999].copy()
df_segunda_conciliacao = marcar_duplicados_com_pior_score(df_segunda_conciliacao)
duplicados = df_segunda_conciliacao[df_segunda_conciliacao["Pontuação"] == 998].copy()
df_segunda_conciliacao = df_segunda_conciliacao[df_segunda_conciliacao["Pontuação"] != 998].copy()
df_terceira_conciliacao = pd.concat([df_terceira_conciliacao, duplicados], ignore_index=True)

display(df_segunda_conciliacao)
display(df_terceira_conciliacao)

# %%
#df_conciliado = pd.concat([df_primeira_conciliacao, df_segunda_conciliacao], ignore_index=True)
#df_conciliado = marcar_duplicados_com_pior_score(df_conciliado)
df_conciliado = df_segunda_conciliacao
df_nao_conciliado = df_terceira_conciliacao
#duplicados = df_conciliado[df_conciliado["Pontuação"] == 998].copy()
#df_conciliado = df_conciliado[df_conciliado["Pontuação"] != 998].copy()
#df_nao_conciliado = pd.concat([df_nao_conciliado, duplicados], ignore_index=True)


# %%
#Marcar na planilha ERP o que já foi usado na conciliação para não ser usado novamente.

def marcar_e_filtrar_chaves_utilizadas(df_erp, df_conciliado):
    """
    Marca chaves já utilizadas no df_erp e retorna um novo DataFrame
    apenas com chaves disponíveis (não utilizadas).
    """

    # Normaliza os valores para garantir comparação precisa
    df_erp["Chave"] = pd.to_numeric(df_erp["Chave"], errors="coerce").astype("Int64")
    df_conciliado["Chave ERP"] = pd.to_numeric(df_conciliado["Chave ERP"], errors="coerce").astype("Int64")

    # Coleta as chaves que já foram utilizadas
    chaves_utilizadas = df_conciliado["Chave ERP"].dropna().unique()
    print(chaves_utilizadas)

    # Marca no df_erp quais foram utilizadas
    df_erp["Usada"] = df_erp["Chave"].isin(chaves_utilizadas)

    # Filtra as que ainda estão disponíveis para nova conciliação
    df_erp_disponivel = df_erp[~df_erp["Usada"]].copy()

    print(f" Total de chaves utilizadas: {df_erp['Usada'].sum()}")
    print(f" Total de chaves disponíveis para nova iteração: {len(df_erp_disponivel)}")

    return df_erp, df_erp_disponivel

df_erp, df_erp_disponivel = marcar_e_filtrar_chaves_utilizadas(df_erp, df_conciliado)

# %%
df_nao_conciliado[["Autorização ERP", "NSU ERP", "Chave ERP", "Valor ERP", "DIF_DIAS", "DIF_VALOR", "Status", "Pontuação"]] = df_nao_conciliado.apply(
    lambda row: selecionar_melhor_por_pontuacao_com_autorizacao_e_nsu(row, df_erp_disponivel, 30, 100000.00, True),
    axis=1
)
display(df_nao_conciliado)

# %%
df_nao_conciliado.to_excel(r"C:\Users\yan.fernandes\Desktop\nao_conciliado.xlsx", index=False)
#df_erp.to_excel(r"C:\Users\regiel\Documents\PythonAppsSantander\ERPMarcado.xlsx", index=False)

# %%
def gerar_relatorio_txt(df_conciliado, df_nao_conciliado, df_cancelamento_venda, valor_aluguel_maquina, nome_arquivo="relatorio_conciliacao.txt"):
    def resumo(df, nome):
        total_liquido = df["VALOR LÍQUIDO"].sum()
        total_parcela = df["VALOR DA PARCELA"].sum()
        qtd_titulos = len(df)

        return (
            f" {nome}\n"
            f"  - Valor Líquido Total: R$ {total_liquido:,.2f}\n"
            f"  - Valor da Parcela Total: R$ {total_parcela:,.2f}\n"
            f"  - Quantidade de Títulos: {qtd_titulos}\n"
        )

    # Cálculo dos totais
    total_conciliado = df_conciliado["VALOR LÍQUIDO"].sum()
    total_nao_conciliado = df_nao_conciliado["VALOR LÍQUIDO"].sum()
    total_cancelado = df_cancelamento_venda["VALOR LÍQUIDO"].sum()

    total_banco = total_conciliado + total_nao_conciliado + total_cancelado + valor_aluguel_maquina

    # Geração do texto do relatório
    relatorio = []
    relatorio.append(" RELATÓRIO DE CONCILIAÇÃO\n")
    relatorio.append("=" * 50 + "\n")
    relatorio.append(resumo(df_conciliado, "CONCILIADO"))
    relatorio.append(resumo(df_nao_conciliado, "NÃO CONCILIADO"))
    relatorio.append(resumo(df_cancelamento_venda, "CANCELAMENTO DE VENDA"))
    relatorio.append("=" * 50 + "\n")
    relatorio.append(f" Valor total de aluguel de máquineta: R$ {valor_aluguel_maquina:,.2f}\n")
    relatorio.append(f" Valor Total no Banco: R$ {total_banco:,.2f}\n")

    # Grava no .txt
    with open(nome_arquivo, "w", encoding="utf-8") as f:
        f.writelines(relatorio)

    print(f" Relatório gerado com sucesso: {nome_arquivo}")

# %%
gerar_relatorio_txt(df_conciliado, df_nao_conciliado, df_cancelamento_venda, valor_aluguel_maquina)


def gerar_relatorio_excel(df_conciliado, df_nao_conciliado, df_cancelamento_venda, valor_aluguel_maquina):
    """
    Gera o relatório em formato de DataFrame para ser salvo como aba no Excel
    """
    total_conciliado = df_conciliado['VALOR LÍQUIDO'].sum()
    total_nao_conciliado = df_nao_conciliado['VALOR LÍQUIDO'].sum()
    total_cancelado = df_cancelamento_venda['VALOR LÍQUIDO'].sum()
    total_banco = total_conciliado + total_nao_conciliado + total_cancelado + valor_aluguel_maquina

    relatorio_data = {
        'Descrição': [
            ' RELATÓRIO DE CONCILIAÇÃO', '',
            ' CONCILIADO',
            '  - Valor Líquido Total',
            '  - Valor da Parcela Total',
            '  - Quantidade de Títulos', '',
            ' NÃO CONCILIADO',
            '  - Valor Líquido Total',
            '  - Valor da Parcela Total',
            '  - Quantidade de Títulos', '',
            ' CANCELAMENTO DE VENDA',
            '  - Valor Líquido Total',
            '  - Valor da Parcela Total',
            '  - Quantidade de Títulos', '',
            ' Valor total de aluguel de máquineta', '',
            ' Valor Total no Banco'
        ],
        'Valor': [
            '', '',
            '',
            f"R$ {total_conciliado:,.2f}",
            f"R$ {df_conciliado['VALOR DA PARCELA'].sum():,.2f}",
            len(df_conciliado), '',
            '',
            f"R$ {total_nao_conciliado:,.2f}",
            f"R$ {df_nao_conciliado['VALOR DA PARCELA'].sum():,.2f}",
            len(df_nao_conciliado), '',
            '',
            f"R$ {total_cancelado:,.2f}",
            f"R$ {df_cancelamento_venda['VALOR DA PARCELA'].sum():,.2f}",
            len(df_cancelamento_venda), '',
            f"R$ {valor_aluguel_maquina:,.2f}", '',
            f"R$ {total_banco:,.2f}"
        ]
    }

    return pd.DataFrame(relatorio_data)

def exportar_consolidado(df_conciliado, df_nao_conciliado, df_cancelamento_venda, 
                        df_aluguel_maquina, valor_aluguel_maquina, caminho_saida):
    """
    Exporta todos os resultados para uma planilha Excel com abas separadas
    """
    with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
        # Salvar cada DataFrame em uma aba diferente
        df_conciliado.to_excel(writer, sheet_name='Conciliados', index=False)
        df_nao_conciliado.to_excel(writer, sheet_name='Não Conciliados', index=False)
        df_cancelamento_venda.to_excel(writer, sheet_name='Cancelamentos', index=False)
        df_aluguel_maquina.to_excel(writer, sheet_name='Aluguel e Tarifas', index=False)
        
        # Adicionar aba com o relatório
        relatorio_df = gerar_relatorio_excel(df_conciliado, df_nao_conciliado, 
                                        df_cancelamento_venda, valor_aluguel_maquina)
        relatorio_df.to_excel(writer, sheet_name='Relatório', index=False, header=False)
        
        # Formatação da aba de relatório
        workbook = writer.book
        worksheet = writer.sheets['Relatório']
        
        # Ajustar largura das colunas
        worksheet.column_dimensions['A'].width = 35
        worksheet.column_dimensions['B'].width = 20
        
        # Aplicar formatação
        bold_font = Font(bold=True)
        for row in [1, 3, 8, 13, 17, 19]:
            worksheet[f'A{row}'].font = bold_font
            worksheet[f'B{row}'].font = bold_font
        
        # Formatar valores monetários
        for row in range(4, 20):
            if worksheet[f'B{row}'].value and 'R$' in str(worksheet[f'B{row}'].value):
                worksheet[f'B{row}'].number_format = '"R$"#,##0.00'

# Substitua as linhas finais do código original (onde estão os to_excel individuais) por:
caminho_final = r"C:\Users\yan.fernandes\Desktop\nao_conciliado.xlsx"
exportar_consolidado(
    df_conciliado=df_conciliado,
    df_nao_conciliado=df_nao_conciliado,
    df_cancelamento_venda=df_cancelamento_venda,
    df_aluguel_maquina=df_aluguel_maquina,
    valor_aluguel_maquina=valor_aluguel_maquina,
    caminho_saida=caminho_final
)

# Mantenha esta linha se quiser continuar gerando o TXT também
gerar_relatorio_txt(df_conciliado, df_nao_conciliado, df_cancelamento_venda, valor_aluguel_maquina)




# %%
# Definir caminho de saída (ajuste conforme necessário)
caminho_final = r"C:\Users\yan.fernandes\Desktop\nao_conciliado.xlsx"
# Chamar a função de exportação consolidada
exportar_consolidado(
    df_conciliado=df_conciliado,
    df_nao_conciliado=df_nao_conciliado,
    df_cancelamento_venda=df_cancelamento_venda,
    df_aluguel_maquina=df_aluguel_maquina,
    valor_aluguel_maquina=valor_aluguel_maquina,
    caminho_saida=caminho_final
)

# Opcional: manter a geração do TXT original
gerar_relatorio_txt(df_conciliado, df_nao_conciliado, df_cancelamento_venda, valor_aluguel_maquina)


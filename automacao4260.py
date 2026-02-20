import os
import pandas as pd
from pathlib import Path

# =====================
# Configurações
# =====================

USUARIO = os.getlogin()
PASTA_ORIGEM = Path(f"C:/Users/{USUARIO}/Desktop/4546_4260")

ARQUIVOS = {
    "4260hsl.xlsx": "Santa Luzia",
    "4260hsh.xlsx": "Santa Helena",
    "4260dfstar.xlsx": "DF Star",
    "4260hcbr.xlsx": "Coração",
}

COLUNAS_REMOVER = ["Telefone", "E-mail"]

COLUNAS_ADICIONAIS = ["Coluna1", "Coluna2", "Coluna3"]

# =====================
# Funções
# =====================

def ajustar_dataframe(df: pd.DataFrame):
    """
    Remove colunas desnecessárias e adiciona colunas extras.
    Reordena as colunas colocando a primeira coluna original na frente.
    """
    df.drop(columns=COLUNAS_REMOVER, inplace=True, errors="ignore")
    for col in COLUNAS_ADICIONAIS:
        df[col] = ""

    primeira_coluna = df.columns[0]
    colunas_reordenadas = [primeira_coluna] + COLUNAS_ADICIONAIS + [
        c for c in df.columns if c not in [primeira_coluna] + COLUNAS_ADICIONAIS
    ]
    return df[colunas_reordenadas]

def aplicar_filtros_unidades(df: pd.DataFrame, unidade: str):
    """
    Ajusta o 'Setor cirurgia' para cada unidade e filtra apenas cirurgias realizadas.
    """
    if unidade == "Santa Luzia":
        df["Setor cirurgia"] = df["Setor cirurgia"].replace("Centro Cirurgico One Day", "Centro Cirúrgico (HSLB)")
        procedimentos = ["Cesariana", "Cesariana (Feto Único Ou Múltiplo)", "Parto (Via Vaginal)", "Parto via Baixa"]
        df.loc[df["Procedimento"].isin(procedimentos), "Setor cirurgia"] = "Centro Obstétrico (HSLB)"
    elif unidade == "Santa Helena":
        df["Setor cirurgia"] = df["Setor cirurgia"].replace("Centro Cirúrgico One Day", "Centro Cirúrgico (HSHB)")
        df["Setor cirurgia"] = df["Setor cirurgia"].replace("Hemodinâmica Exames (HCBR)", "Hemodinâmica (HSHB)")
        procedimentos = [
            "Cesariana",
            "Cesariana (Feto Único Ou Múltiplo)",
            "Parto (Via Vaginal)",
            "Parto Gemelar (Cada um Subsequente ao Parto)",
            "Parto Múltiplo Por Via Vaginal (Cada Um Subsequente Ao Inicial)",
            "Parto via Baixa",
        ]
        df.loc[df["Procedimento"].isin(procedimentos), "Setor cirurgia"] = "Centro Obstétrico (HSHB)"
        df.loc[~df["Setor cirurgia"].isin(["Hemodinâmica (HSHB)"]), "Setor cirurgia"] = "Centro Cirúrgico (HSHB)"
    elif unidade == "DF Star":
        df["Setor cirurgia"] = df["Setor cirurgia"].replace("RPA Centro Cirúrgico (DFStar)", "Centro Cirúrgico (DFStar)")
    elif unidade == "Coração":
        df["Setor cirurgia"] = df["Setor cirurgia"].replace("RPA Hemodinâmica (HCBR)", "Hemodinâmica (HCBR) - CIC")

    df = df[df.get("Status cirurgia", df.get("Status", None)) == "Realizada"]
    return df

# =====================
# Execução
# =====================

if __name__ == "__main__":
    for arquivo, unidade in ARQUIVOS.items():
        caminho_arquivo = PASTA_ORIGEM / arquivo
        df = pd.read_excel(caminho_arquivo)
        df = ajustar_dataframe(df)
        df = aplicar_filtros_unidades(df, unidade)

        nome_saida = f"T{arquivo}"
        df.to_excel(PASTA_ORIGEM / nome_saida, sheet_name=f"{arquivo.split('.')[0]} {unidade}", index=False)
        print(f"✅ {nome_saida} processado e salvo para {unidade}")
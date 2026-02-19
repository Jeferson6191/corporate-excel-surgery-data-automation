import os
import pandas as pd
from pathlib import Path

# =====================
# Configurações
# =====================

USUARIO = os.getlogin()
PASTA_ORIGEM = Path(f"C:/Users/{USUARIO}/Desktop/4546_4260")

ARQUIVOS = [
    "4546hsh.xlsx",
    "4546hsl.xlsx",
]

SETOR_POR_ARQUIVO = {
    "4546hsh.xlsx": [
        "Centro Obstétrico (HSHB)",
        "Centro Cirúrgico (HSHB)",
        "Hemodinâmica (HSHB)",
    ],
    "4546hsl.xlsx": [
        "Centro Obstétrico (HSLB)",
        "Centro Cirúrgico (HSLB)",
        "3",
    ],
}

PROCEDIMENTOS_PARTO = [
    "Parto (Via Vaginal)",
    "Parto via Baixa",
    "Cesariana (Feto Único Ou Múltiplo)",
    "Cesariana",
]

COLUNAS_REMOVER = [
    "Ds. diagn. pós",
    "Ds. diagóstico",
    "Aval Pré Anestésica",
]

# =====================
# Funções
# =====================

def aplicar_filtros(df: pd.DataFrame, arquivo: str):
    setor_obst, setor_cir, setor_extra = SETOR_POR_ARQUIVO[arquivo]

    filtro_partos = (
        (df["Tipo"] == "Cirurgia") &
        (df["Status"] == "Realizada") &
        (df["Procedimento"].isin(PROCEDIMENTOS_PARTO))
    )

    filtro_cirurgico = (
        (df["Tipo"] == "Cirurgia") &
        (df["Status"] == "Realizada") &
        (df["Setor cirurgia"] == setor_obst) &
        (~df["Procedimento"].isin(PROCEDIMENTOS_PARTO))
    )

    filtro_hemodinamica = (
        (df["Tipo"] == "Cirurgia") &
        (df["Status"] == "Realizada") &
        (df["Setor cirurgia"].str.startswith("Hemodin", na=False))
    )

    filtro_outros = (
        (df["Tipo"] == "Cirurgia") &
        (df["Status"] == "Realizada") &
        (~df["Setor cirurgia"].isin([setor_obst, setor_cir, setor_extra]))
    )

    df.loc[filtro_partos, "Setor cirurgia"] = setor_obst
    df.loc[filtro_cirurgico, "Setor cirurgia"] = setor_cir

    if arquivo == "4546hsh.xlsx":
        df.loc[filtro_hemodinamica, "Setor cirurgia"] = "Hemodinâmica (HSHB)"

    df.loc[filtro_outros, "Setor cirurgia"] = setor_cir

    return df


def processar_arquivo(nome_arquivo: str):
    caminho = PASTA_ORIGEM / nome_arquivo
    df = pd.read_excel(caminho)

    df = aplicar_filtros(df, nome_arquivo)

    df = df.drop(columns=COLUNAS_REMOVER, errors="ignore")

    saida = PASTA_ORIGEM / f"T{nome_arquivo}"
    df.to_excel(saida, index=False)


# =====================
# Execução
# =====================

if __name__ == "__main__":
    for arquivo in ARQUIVOS:
        processar_arquivo(arquivo)

import os
import pandas as pd
from pathlib import Path
import subprocess

# =====================
# Configurações
# =====================

USUARIO = os.getlogin()
PASTA_ORIGEM = Path(f"C:/Users/{USUARIO}/Desktop/4546_4260")

ARQUIVOS = [
    "4546hsh.xlsx",
    "4546hsl.xlsx",
    "4546dfstar.xlsx",
    "4546hcbr.xlsx",
]

SETOR_POR_ARQUIVO = {
    "4546hsh.xlsx": ["Centro Obstétrico (HSHB)", "Centro Cirúrgico (HSHB)", "Hemodinâmica (HSHB)"],
    "4546hsl.xlsx": ["Centro Obstétrico (HSLB)", "Centro Cirúrgico (HSLB)", "3"],
}

PROCEDIMENTOS_PARTO = [
    "Parto (Via Vaginal)",
    "Parto via Baixa",
    "Cesariana (Feto Único Ou Múltiplo)",
    "Cesariana",
]

COLUNAS_REMOVER = ["Ds. diagn. pós", "Ds. diagóstico", "Aval Pré Anestésica"]

# =====================
# Funções
# =====================

def aplicar_filtros(df: pd.DataFrame, arquivo: str):
    """
    Ajusta a coluna 'Setor cirurgia' para os arquivos HSH e HSL,
    aplicando filtros de procedimentos obstétricos e setores específicos.
    """
    setor_obst, setor_cir, setor_extra = SETOR_POR_ARQUIVO.get(arquivo, (None, None, None))

    if setor_obst and setor_cir:
        filtro_partos = (
            (df["Tipo"] == "Cirurgia")
            & (df["Status"] == "Realizada")
            & (df["Procedimento"].isin(PROCEDIMENTOS_PARTO))
        )

        filtro_cirurgico = (
            (df["Tipo"] == "Cirurgia")
            & (df["Status"] == "Realizada")
            & (df["Setor cirurgia"] == setor_obst)
            & (~df["Procedimento"].isin(PROCEDIMENTOS_PARTO))
        )

        filtro_hemodinamica = (
            (df["Tipo"] == "Cirurgia")
            & (df["Status"] == "Realizada")
            & (df["Setor cirurgia"].str.startswith("Hemodin", na=False))
        )

        filtro_outros = (
            (df["Tipo"] == "Cirurgia")
            & (df["Status"] == "Realizada")
            & (~df["Setor cirurgia"].isin([setor_obst, setor_cir, setor_extra]))
        )

        df.loc[filtro_partos, "Setor cirurgia"] = setor_obst
        df.loc[filtro_cirurgico, "Setor cirurgia"] = setor_cir
        if arquivo == "4546hsh.xlsx":
            df.loc[filtro_hemodinamica, "Setor cirurgia"] = "Hemodinâmica (HSHB)"
        df.loc[filtro_outros, "Setor cirurgia"] = setor_cir

    return df

def aplicar_filtros_adicionais(df: pd.DataFrame, arquivo: str):
    """
    Aplica filtros para arquivos adicionais DFStar e HCBR.
    """
    if arquivo == "4546dfstar.xlsx":
        filtro = df["Setor cirurgia"].str.startswith("RPA Centro Cirúrgico (DFStar)", na=False)
        df.loc[filtro, "Setor cirurgia"] = "Centro Cirúrgico (DFStar)"
    elif arquivo == "4546hcbr.xlsx":
        filtro = df["Setor cirurgia"].str.startswith("RPA Hemodinâmica (HCBR)", na=False)
        df.loc[filtro, "Setor cirurgia"] = "Hemodinâmica (HCBR) - CIC"
    return df

def processar_arquivo(nome_arquivo: str):
    """
    Lê, processa e salva o arquivo Excel aplicando todos os filtros necessários.
    """
    caminho = PASTA_ORIGEM / nome_arquivo
    df = pd.read_excel(caminho)

    if nome_arquivo in SETOR_POR_ARQUIVO:
        df = aplicar_filtros(df, nome_arquivo)

    df = aplicar_filtros_adicionais(df, nome_arquivo)
    df = df.drop(columns=COLUNAS_REMOVER, errors="ignore")

    saida = PASTA_ORIGEM / f"T{nome_arquivo}"
    df.to_excel(saida, index=False)
    print(f"{saida.name} processado e salvo.")

# =====================
# Execução
# =====================

if __name__ == "__main__":
    for arquivo in ARQUIVOS:
        processar_arquivo(arquivo)

    # Executa o subprocesso
    caminho_subprocesso = Path(f"C:/Users/{USUARIO}/Desktop/testes/automacao4260.py")
    try:
        subprocess.run(["python", str(caminho_subprocesso)], check=True)
        print("✅ Subprocesso automacao4260.py executado com sucesso.")
    except subprocess.CalledProcessError:
        print("❌ Erro ao executar o subprocesso automacao4260.py.")
import pandas as pd
import os


ARQUIVO_ENTRADA = r"C:\Users\USUARIO\Downloads\custos_ti_2024.xlsx"
ARQUIVO_SAIDA = r"C:\Users\USUARIO\Desktop\custos_ti_2024_tratado.xlsx"

def carregar_dados(caminho):
    print(f"Carregando arquivo: {caminho}")
    df = pd.read_excel(caminho, sheet_name="Lançamentos")
    print(f"  {len(df)} linhas carregadas | {len(df.columns)} colunas")
    return df


def verificar_nulos(df):
    print("\nVerificando valores nulos...")
    nulos = df.isnull().sum()
    nulos = nulos[nulos > 0]
    if nulos.empty:
        print("  Nenhum valor nulo encontrado.")
    else:
        print(f"  Valores nulos encontrados:\n{nulos}")
    return df.dropna()


def corrigir_tipos(df):
    print("\nCorrigindo tipos de dados...")
    df["Data"] = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")
    df["Mês"] = df["Mês"].astype(int)
    df["Quantidade"] = df["Quantidade"].astype(int)
    df["Valor Unitário (R$)"] = pd.to_numeric(df["Valor Unitário (R$)"], errors="coerce").round(2)
    df["Valor Total (R$)"] = pd.to_numeric(df["Valor Total (R$)"], errors="coerce").round(2)
    print("  Tipos corrigidos com sucesso.")
    return df


def padronizar_texto(df):
    print("\nPadronizando campos de texto...")
    colunas_texto = ["Categoria", "Descrição", "Fornecedor", "Departamento", "Status", "Mês Nome", "Trimestre"]
    for col in colunas_texto:
        if col in df.columns:
            df[col] = df[col].str.strip().str.title()
    print("  Texto padronizado com sucesso.")
    return df


def validar_valores(df):
    print("\nValidando valores financeiros...")
    negativos = df[df["Valor Total (R$)"] < 0]
    if not negativos.empty:
        print(f"  Atenção: {len(negativos)} registros com valor negativo encontrados e removidos.")
        df = df[df["Valor Total (R$)"] >= 0]
    else:
        print("  Nenhum valor negativo encontrado.")
    return df


def remover_duplicatas(df):
    print("\nVerificando duplicatas...")
    antes = len(df)
    df = df.drop_duplicates(subset=["ID Lançamento"])
    depois = len(df)
    removidos = antes - depois
    if removidos > 0:
        print(f"  {removidos} registros duplicados removidos.")
    else:
        print("  Nenhuma duplicata encontrada.")
    return df


def exportar(df, caminho):
    print(f"\nExportando arquivo tratado: {caminho}")
    df.to_excel(caminho, index=False, sheet_name="Lançamentos Tratados")
    print(f"  Arquivo salvo com sucesso! ({len(df)} linhas)")


def resumo_final(df):
    print("\n" + "="*45)
    print("RESUMO DO TRATAMENTO")
    print("="*45)
    print(f"Total de registros: {len(df)}")
    print(f"Período: {df['Data'].min().strftime('%d/%m/%Y')} a {df['Data'].max().strftime('%d/%m/%Y')}")
    print(f"Total gasto: R$ {df['Valor Total (R$)'].sum():,.2f}")
    print("\nGastos por categoria:")
    resumo = df.groupby("Categoria")["Valor Total (R$)"].sum().sort_values(ascending=False)
    for cat, valor in resumo.items():
        print(f"  {cat}: R$ {valor:,.2f}")
    print("="*45)


def main():
    if not os.path.exists(ARQUIVO_ENTRADA):
        print(f"Erro: arquivo '{ARQUIVO_ENTRADA}' não encontrado.")
        print("Certifique-se de que o arquivo está na mesma pasta que este script.")
        return

    df = carregar_dados(ARQUIVO_ENTRADA)
    df = verificar_nulos(df)
    df = corrigir_tipos(df)
    df = padronizar_texto(df)
    df = validar_valores(df)
    df = remover_duplicatas(df)
    exportar(df, ARQUIVO_SAIDA)
    resumo_final(df)


if __name__ == "__main__":
    main()

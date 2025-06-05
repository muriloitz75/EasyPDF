import pandas as pd
import openpyxl
from openpyxl import load_workbook
from pdf_extractor import extract_data_from_pdf

# Extrair dados diretamente do PDF
print("Extraindo dados do PDF original...")
pdf_data = extract_data_from_pdf('input/2025.pdf')
pdf_df = pd.DataFrame(pdf_data)

# Verificar algumas linhas para entender como o ISS próprio e o ISS retido estão sendo extraídos
print("\nPrimeiras 5 linhas extraídas diretamente do PDF:")
print(pdf_df[['Número da Nota', 'ISS Próprio (R$)', 'ISS Retido (R$)']].head())

# Verificar linhas onde ISS próprio e ISS retido são iguais
mask = pdf_df['ISS Próprio (R$)'] == pdf_df['ISS Retido (R$)']
equal_values = pdf_df[mask]
print(f"\nLinhas onde ISS Próprio = ISS Retido no PDF original: {len(equal_values)}")
if len(equal_values) > 0:
    print(equal_values[['Número da Nota', 'ISS Próprio (R$)', 'ISS Retido (R$)']].head())

# Carregar o Excel gerado
print("\nVerificando o Excel gerado...")
excel_df = pd.read_excel('output/teste_sem_total_iss.xlsx', sheet_name='Dados_NFSe')

# Verificar algumas linhas para entender como o ISS próprio e o ISS retido estão sendo exibidos
print("\nPrimeiras 5 linhas do Excel gerado:")
print(excel_df[['Número da Nota', 'ISS Próprio (R$)', 'ISS Retido (R$)']].head())

# Verificar linhas onde ISS próprio e ISS retido são iguais
mask = excel_df['ISS Próprio (R$)'] == excel_df['ISS Retido (R$)']
equal_values = excel_df[mask]
print(f"\nLinhas onde ISS Próprio = ISS Retido no Excel gerado: {len(equal_values)}")
if len(equal_values) > 0:
    print(equal_values[['Número da Nota', 'ISS Próprio (R$)', 'ISS Retido (R$)']].head())

# Verificar se há diferenças entre o PDF e o Excel
print("\nVerificando diferenças entre PDF e Excel...")
merged_df = pd.merge(
    pdf_df[['Número da Nota', 'ISS Próprio (R$)', 'ISS Retido (R$)']],
    excel_df[['Número da Nota', 'ISS Próprio (R$)', 'ISS Retido (R$)']],
    on='Número da Nota',
    suffixes=('_pdf', '_excel')
)

# Verificar diferenças no ISS Próprio
iss_proprio_diff = merged_df[abs(merged_df['ISS Próprio (R$)_pdf'] - merged_df['ISS Próprio (R$)_excel']) > 0.01]
print(f"\nLinhas com diferença no ISS Próprio: {len(iss_proprio_diff)}")
if len(iss_proprio_diff) > 0:
    print(iss_proprio_diff[['Número da Nota', 'ISS Próprio (R$)_pdf', 'ISS Próprio (R$)_excel']].head())

# Verificar diferenças no ISS Retido
iss_retido_diff = merged_df[abs(merged_df['ISS Retido (R$)_pdf'] - merged_df['ISS Retido (R$)_excel']) > 0.01]
print(f"\nLinhas com diferença no ISS Retido: {len(iss_retido_diff)}")
if len(iss_retido_diff) > 0:
    print(iss_retido_diff[['Número da Nota', 'ISS Retido (R$)_pdf', 'ISS Retido (R$)_excel']].head())

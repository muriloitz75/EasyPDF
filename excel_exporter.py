import pandas as pd
from datetime import datetime
import logging
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# --- Configuração de Logging ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', handlers=[logging.FileHandler("excel_exporter.log"), logging.StreamHandler()])
logger = logging.getLogger("excel_exporter")

# --- Constantes de Estilo ---
TITLE_FONT = Font(bold=True, size=16, color="FFFFFF")
SECTION_FONT = Font(bold=True, size=14, color="FFFFFF")
KPI_TITLE_FONT = Font(bold=True, size=11)
KPI_VALUE_FONT = Font(size=18, bold=True, color="2F5496")
TABLE_HEADER_FONT = Font(bold=True, color="FFFFFF")

TITLE_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
SECTION_FILL = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
TABLE_HEADER_FILL = PatternFill(start_color="80A6D9", end_color="80A6D9", fill_type="solid")
KPI_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")

CENTER_ALIGN = Alignment(horizontal='center', vertical='center')
RIGHT_ALIGN = Alignment(horizontal='right', vertical='center')
LEFT_ALIGN = Alignment(horizontal='left', vertical='center')

THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

# --- Funções Principais de Exportação ---

def export_to_excel(df, output_path=None):
    if df.empty:
        logger.warning("DataFrame vazio. Nenhum arquivo Excel gerado.")
        return None

    # ... (lógica de preparação do df principal, como correção de 'Incidência', etc.)

    if output_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = f"notas_fiscais_{timestamp}.xlsx"

    logger.info(f"Iniciando exportação para Excel: {output_path}")
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Dados_NFSe', index=False)
            # format_excel_sheet(writer, 'Dados_NFSe', df) # A formatação detalhada pode ser adicionada aqui
            create_summary_dashboard(writer, df, sheet_name='Resumo_NFSe')
        logger.info(f"Exportação para Excel concluída: {output_path}")
        return output_path
    except Exception as e:
        logger.error(f"Erro ao exportar para Excel: {str(e)}")
        raise

def create_summary_dashboard(writer, df, sheet_name):
    logger.info(f"Criando dashboard de resumo: {sheet_name}")
    if df.empty:
        return

    if 'Competência' not in df.columns:
        if 'Mês' in df.columns and 'Ano' in df.columns:
            df['Competência'] = df.apply(lambda x: f"{int(x['Mês']):02d}/{int(x['Ano'])}" if pd.notna(x['Mês']) and pd.notna(x['Ano']) else "Desconhecida", axis=1)
        else:
            logger.error("Colunas 'Mês' e 'Ano' ausentes para criar a competência.")
            return

    worksheet = writer.book.create_sheet(sheet_name)
    worksheet.sheet_view.showGridLines = False # Remove as linhas de grade para um visual mais limpo

    # Título Geral
    worksheet.merge_cells('A1:L1')
    title_cell = worksheet['A1']
    title_cell.value = "DASHBOARD DE ANÁLISE DE NOTAS FISCAIS DE SERVIÇO"
    title_cell.font = TITLE_FONT
    title_cell.alignment = CENTER_ALIGN
    title_cell.fill = TITLE_FILL
    worksheet.row_dimensions[1].height = 30
    
    current_row = 3

    try:
        df['temp_date'] = pd.to_datetime(df['Competência'], format='%m/%Y', errors='coerce')
        competencias = sorted(df['Competência'].dropna().unique(), key=lambda x: datetime.strptime(x, "%m/%Y"))
        df.drop('temp_date', axis=1, inplace=True)
    except Exception:
        competencias = sorted(df['Competência'].dropna().unique())

    for competencia in competencias:
        df_competencia = df[df['Competência'] == competencia].copy()
        if df_competencia.empty:
            continue

        # Cabeçalho da Seção da Competência
        worksheet.merge_cells(f'A{current_row}:L{current_row}')
        section_title = worksheet[f'A{current_row}']
        section_title.value = f"Análise da Competência: {competencia}"
        section_title.font = SECTION_FONT
        section_title.alignment = CENTER_ALIGN
        section_title.fill = SECTION_FILL
        worksheet.row_dimensions[current_row].height = 25
        current_row += 1

        # Bloco de KPIs
        current_row = _create_kpi_block(worksheet, df_competencia, current_row)
        
        # Bloco de Tabelas
        current_row = _create_tables_block(worksheet, df_competencia, current_row)

        current_row += 2 # Espaço entre os blocos de competência

    # --- Bloco de Total Geral ---
    current_row += 1 # Espaço extra antes do total geral
    worksheet.merge_cells(f'A{current_row}:L{current_row}')
    total_title = worksheet[f'A{current_row}']
    total_title.value = "TOTAL GERAL DO PERÍODO"
    total_title.font = Font(bold=True, size=16, color="FFFFFF")
    total_title.alignment = CENTER_ALIGN
    total_title.fill = TITLE_FILL # Usando a cor do título principal para mais destaque
    worksheet.row_dimensions[current_row].height = 30
    current_row += 1

    # KPIs e Tabelas para o DataFrame completo
    current_row = _create_kpi_block(worksheet, df, current_row)
    _create_tables_block(worksheet, df, current_row)

def _create_kpi_block(worksheet, df, start_row):
    df_validas = df[df['Situação'] != 'CANCELADA']
    
    total_faturado = df_validas['Valor do Serviço (R$)'].sum()
    qtd_notas = len(df_validas)
    iss_total = df_validas['ISS Próprio (R$)'].sum() + df_validas['ISS Retido (R$)'].sum()
    ticket_medio = total_faturado / qtd_notas if qtd_notas > 0 else 0

    kpis = {
        "Valor Total Faturado": f"R$ {total_faturado:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
        "Quantidade de Notas Válidas": f"{qtd_notas}",
        "ISS Total (Próprio + Retido)": f"R$ {iss_total:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'),
        "Ticket Médio por Nota": f"R$ {ticket_medio:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    }

    col = 1
    for title, value in kpis.items():
        # Merge 3 colunas para cada KPI
        worksheet.merge_cells(start_row=start_row, start_column=col, end_row=start_row, end_column=col + 2)
        title_cell = worksheet.cell(row=start_row, column=col, value=title)
        title_cell.font = KPI_TITLE_FONT
        title_cell.alignment = CENTER_ALIGN
        title_cell.fill = KPI_FILL
        title_cell.border = THIN_BORDER

        worksheet.merge_cells(start_row=start_row + 1, start_column=col, end_row=start_row + 1, end_column=col + 2)
        value_cell = worksheet.cell(row=start_row + 1, column=col, value=value)
        value_cell.font = KPI_VALUE_FONT
        value_cell.alignment = CENTER_ALIGN
        value_cell.border = THIN_BORDER
        
        # Formatação das células mescladas
        for c in range(col, col + 3):
            worksheet.cell(row=start_row, column=c).border = THIN_BORDER
            worksheet.cell(row=start_row + 1, column=c).border = THIN_BORDER

        col += 3 # Pula para o próximo bloco de 3 colunas

    worksheet.row_dimensions[start_row].height = 20
    worksheet.row_dimensions[start_row + 1].height = 30
    
    return start_row + 3

def _create_tables_block(worksheet, df, start_row):
    # Tabela 1: Resumo por Situação (colunas A-E)
    row_after_situacao = _create_situacao_summary(worksheet, df, start_row, start_col=1)
    
    # Tabela 2: Resumo por Código de Serviço (colunas G-L)
    row_after_codigo = _create_codigo_servico_summary(worksheet, df, start_row, start_col=7)

    # A próxima linha será a maior entre as duas tabelas
    next_row = max(row_after_situacao, row_after_codigo)

    # Tabela 3: Top Tomadores
    next_row = _create_top_tomadores_summary(worksheet, df, next_row, start_col=1)

    return next_row

def _write_table_to_sheet(worksheet, df_summary, headers, start_row, start_col, column_widths):
    # Escreve cabeçalhos
    for i, header in enumerate(headers):
        cell = worksheet.cell(row=start_row, column=start_col + i, value=header)
        cell.font = TABLE_HEADER_FONT
        cell.fill = TABLE_HEADER_FILL
        cell.alignment = CENTER_ALIGN
        cell.border = THIN_BORDER
        worksheet.column_dimensions[get_column_letter(start_col + i)].width = column_widths[i]

    # Escreve dados
    current_row = start_row + 1
    for _, row_data in df_summary.iterrows():
        for i, col_name in enumerate(headers):
            # Formatação de valores monetários para o padrão brasileiro
            if "R$" in col_name and isinstance(row_data[col_name], (int, float)):
                # Converter para o formato brasileiro (vírgula como decimal, ponto como milhar)
                value = f"{row_data[col_name]:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                cell = worksheet.cell(row=current_row, column=start_col + i, value=value)
            else:
                cell = worksheet.cell(row=current_row, column=start_col + i, value=row_data[col_name])
                
            cell.border = THIN_BORDER
            # Formatação específica
            if "R$" in col_name:
                cell.number_format = 'R$ #.##0,00'
                cell.alignment = RIGHT_ALIGN
            elif "%" in col_name:
                cell.number_format = '0.00"%"'
                cell.alignment = CENTER_ALIGN
            elif "Qtd." in col_name:
                cell.alignment = CENTER_ALIGN
            else:
                cell.alignment = LEFT_ALIGN
            
            # Formatação de linha de total
            if str(row_data.get(headers[0])).upper() == 'TOTAL':
                cell.font = Font(bold=True)
                cell.fill = KPI_FILL

        current_row += 1
    return current_row

def _create_situacao_summary(worksheet, df, start_row, start_col):
    summary = df.groupby('Situação').agg(
        Qtd_Notas=('Número da Nota', 'count'),
        Valor_Servico=('Valor do Serviço (R$)', 'sum')
    ).reset_index().rename(columns={'Valor_Servico': 'Valor Serviço (R$)'})
    total_row = summary.sum(numeric_only=True)
    total_row['Situação'] = 'TOTAL'
    summary = pd.concat([summary, pd.DataFrame([total_row])], ignore_index=True)
    
    headers = ['Situação', 'Qtd_Notas', 'Valor Serviço (R$)']
    widths = [18, 12, 20]
    
    return _write_table_to_sheet(worksheet, summary, headers, start_row, start_col, widths)

def _create_codigo_servico_summary(worksheet, df, start_row, start_col):
    df_validas = df[df['Situação'] != 'CANCELADA'].copy()
    summary = df_validas.groupby('Código do Serviço').agg(
        Qtd_Notas=('Número da Nota', 'count'),
        Valor_Servico=('Valor do Serviço (R$)', 'sum'),
        Aliquota_perc=('Alíquota (%)', 'first')
    ).reset_index().rename(columns={'Valor_Servico': 'Valor Serviço (R$)', 'Aliquota_perc': 'Alíquota (%)'})
    total_row = summary.sum(numeric_only=True)
    total_row['Código do Serviço'] = 'TOTAL'
    summary = pd.concat([summary, pd.DataFrame([total_row])], ignore_index=True)

    headers = ['Código do Serviço', 'Qtd_Notas', 'Valor Serviço (R$)', 'Alíquota (%)']
    widths = [18, 12, 20, 12]

    return _write_table_to_sheet(worksheet, summary, headers, start_row, start_col, widths)

def _create_top_tomadores_summary(worksheet, df, start_row, start_col):
    df_validas = df[df['Situação'] != 'CANCELADA'].copy()
    summary = df_validas.groupby('Nome do Tomador').agg(
        Valor_Total=('Valor do Serviço (R$)', 'sum'),
        Qtd_Notas=('Número da Nota', 'count')
    ).reset_index()
    summary = summary.sort_values('Valor_Total', ascending=False).head(5)
    summary = summary.rename(columns={'Valor_Total': 'Valor Total (R$)'})

    headers = ['Nome do Tomador', 'Qtd_Notas', 'Valor Total (R$)']
    widths = [40, 12, 20]
    
    # Título da tabela
    worksheet.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=start_col + len(widths) - 1)
    title_cell = worksheet.cell(row=start_row, column=start_col, value="Top 5 Tomadores por Valor")
    title_cell.font = Font(bold=True)
    title_cell.alignment = CENTER_ALIGN
    
    return _write_table_to_sheet(worksheet, summary, headers, start_row + 1, start_col, widths)

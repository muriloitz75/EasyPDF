import pandas as pd
from datetime import datetime
import logging
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Series, Reference, PieChart
from openpyxl.chart.label import DataLabelList

# --- Configuração de Logging ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', handlers=[logging.FileHandler("excel_exporter.log"), logging.StreamHandler()])
logger = logging.getLogger("excel_exporter")

# --- Constantes de Estilo ---
TITLE_FONT = Font(bold=True, size=18, color="FFFFFF")
YEAR_FONT = Font(bold=True, size=16, color="FFFFFF")
SECTION_FONT = Font(bold=True, size=14, color="FFFFFF")
KPI_TITLE_FONT = Font(bold=True, size=11)
KPI_VALUE_FONT = Font(size=18, bold=True, color="2F5496")
TABLE_HEADER_FONT = Font(bold=True, color="FFFFFF")
CHART_TITLE_FONT = Font(bold=True, size=12)

TITLE_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
YEAR_FILL = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
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

    if output_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = f"notas_fiscais_{timestamp}.xlsx"

    logger.info(f"Iniciando exportação para Excel: {output_path}")
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Dados_NFSe', index=False)
            create_summary_dashboard(writer, df, sheet_name='Resumo_NFSe')
            writer.book.move_sheet('Resumo_NFSe', offset=-len(writer.book.sheetnames)+1)
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
            
    df['Ano'] = pd.to_datetime(df['Competência'], format='%m/%Y', errors='coerce').dt.year
    df['Ano'] = df['Ano'].fillna(0).astype(int)

    worksheet = writer.book.create_sheet(sheet_name)
    worksheet.sheet_view.showGridLines = False

    worksheet.merge_cells('A1:P1')
    title_cell = worksheet['A1']
    title_cell.value = "DASHBOARD DE ANÁLISE DE NOTAS FISCAIS DE SERVIÇO"
    title_cell.font = TITLE_FONT
    title_cell.alignment = CENTER_ALIGN
    title_cell.fill = TITLE_FILL
    worksheet.row_dimensions[1].height = 40
    
    current_row = 3
    
    anos = sorted([ano for ano in df['Ano'].unique() if ano != 0])

    for ano in anos:
        df_ano = df[df['Ano'] == ano]
        
        worksheet.merge_cells(f'A{current_row}:P{current_row}')
        year_title = worksheet[f'A{current_row}']
        year_title.value = f"ANÁLISE DO ANO DE {ano}"
        year_title.font = YEAR_FONT
        year_title.alignment = CENTER_ALIGN
        year_title.fill = YEAR_FILL
        worksheet.row_dimensions[current_row].height = 30
        current_row += 2

        competencias = sorted(df_ano['Competência'].dropna().unique(), key=lambda x: datetime.strptime(x, "%m/%Y"))

        for competencia in competencias:
            df_competencia = df_ano[df_ano['Competência'] == competencia].copy()
            if df_competencia.empty:
                continue

            worksheet.merge_cells(f'A{current_row}:P{current_row}')
            section_title = worksheet[f'A{current_row}']
            section_title.value = f"Análise da Competência: {competencia}"
            section_title.font = SECTION_FONT
            section_title.alignment = CENTER_ALIGN
            section_title.fill = SECTION_FILL
            worksheet.row_dimensions[current_row].height = 25
            current_row += 1

            current_row = _create_kpi_block(worksheet, df_competencia, current_row)
            current_row = _create_tables_block(worksheet, df_competencia, current_row, add_charts=False)
            current_row += 2

        worksheet.merge_cells(f'A{current_row}:P{current_row}')
        total_anual_title = worksheet[f'A{current_row}']
        total_anual_title.value = f"TOTAL CONSOLIDADO DO ANO DE {ano}"
        total_anual_title.font = SECTION_FONT
        total_anual_title.alignment = CENTER_ALIGN
        total_anual_title.fill = SECTION_FILL
        worksheet.row_dimensions[current_row].height = 25
        current_row += 1

        current_row = _create_kpi_block(worksheet, df_ano, current_row)
        current_row = _create_tables_block(worksheet, df_ano, current_row, add_charts=True, chart_title_prefix=f"Ano {ano} - ")
        current_row += 3

    worksheet.merge_cells(f'A{current_row}:P{current_row}')
    total_title = worksheet[f'A{current_row}']
    total_title.value = "TOTAL GERAL DO PERÍODO"
    total_title.font = YEAR_FONT
    total_title.alignment = CENTER_ALIGN
    total_title.fill = YEAR_FILL
    worksheet.row_dimensions[current_row].height = 30
    current_row += 2

    current_row = _create_kpi_block(worksheet, df, current_row)
    _create_tables_block(worksheet, df, current_row, add_charts=True, chart_title_prefix="Geral - ")

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
        worksheet.merge_cells(start_row=start_row, start_column=col, end_row=start_row, end_column=col + 3)
        title_cell = worksheet.cell(row=start_row, column=col, value=title)
        title_cell.font = KPI_TITLE_FONT
        title_cell.alignment = CENTER_ALIGN
        title_cell.fill = KPI_FILL
        title_cell.border = THIN_BORDER

        worksheet.merge_cells(start_row=start_row + 1, start_column=col, end_row=start_row + 1, end_column=col + 3)
        value_cell = worksheet.cell(row=start_row + 1, column=col, value=value)
        value_cell.font = KPI_VALUE_FONT
        value_cell.alignment = CENTER_ALIGN
        value_cell.border = THIN_BORDER
        
        for c in range(col, col + 4):
            worksheet.cell(row=start_row, column=c).border = THIN_BORDER
            worksheet.cell(row=start_row + 1, column=c).border = THIN_BORDER

        col += 4
    
    worksheet.row_dimensions[start_row].height = 20
    worksheet.row_dimensions[start_row + 1].height = 30
    
    return start_row + 3

def _create_tables_block(worksheet, df, start_row, add_charts=False, chart_title_prefix=""):
    # Tabela 1: Resumo por Situação
    situacao_start_col = 1
    summary_situacao, situacao_headers, situacao_widths = _create_situacao_summary(df)
    row_after_situacao = _write_table_to_sheet(worksheet, summary_situacao, situacao_headers, start_row, situacao_start_col, situacao_widths)
    
    # Tabela 2: Resumo por Código de Serviço
    codigo_start_col = 8
    summary_codigo, codigo_headers, codigo_widths = _create_codigo_servico_summary(df)
    row_after_codigo = _write_table_to_sheet(worksheet, summary_codigo, codigo_headers, start_row, codigo_start_col, codigo_widths)

    next_row = max(row_after_situacao, row_after_codigo) + 1

    # Tabela 3: Top Tomadores
    tomadores_start_col = 1
    summary_tomadores, tomadores_headers, tomadores_widths = _create_top_tomadores_summary(df)
    
    # Título da tabela
    worksheet.merge_cells(start_row=next_row, start_column=tomadores_start_col, end_row=next_row, end_column=tomadores_start_col + len(tomadores_widths) - 1)
    title_cell = worksheet.cell(row=next_row, column=tomadores_start_col, value=f"{chart_title_prefix}Top 5 Tomadores por Valor")
    title_cell.font = CHART_TITLE_FONT
    title_cell.alignment = CENTER_ALIGN
    
    tomadores_table_start_row = next_row + 1
    row_after_tomadores = _write_table_to_sheet(worksheet, summary_tomadores, tomadores_headers, tomadores_table_start_row, tomadores_start_col, tomadores_widths)

    if add_charts:
        chart_start_row = row_after_tomadores + 1
        
        # Gráfico de Situação
        _create_situacao_pie_chart(worksheet, summary_situacao, "A" + str(chart_start_row), f"{chart_title_prefix}Distribuição por Situação", start_row, situacao_start_col)
        
        # Gráfico de Tomadores
        _create_top_tomadores_bar_chart(worksheet, summary_tomadores, "H" + str(chart_start_row), f"{chart_title_prefix}Top 5 Tomadores", tomadores_table_start_row, tomadores_start_col)
        
        next_row = chart_start_row + 16 # Espaço para os gráficos
    else:
        next_row = row_after_tomadores

    return next_row

def _write_table_to_sheet(worksheet, df_summary, headers, start_row, start_col, column_widths):
    for i, header in enumerate(headers):
        cell = worksheet.cell(row=start_row, column=start_col + i, value=header)
        cell.font = TABLE_HEADER_FONT
        cell.fill = TABLE_HEADER_FILL
        cell.alignment = CENTER_ALIGN
        cell.border = THIN_BORDER
        if len(column_widths) > i:
            worksheet.column_dimensions[get_column_letter(start_col + i)].width = column_widths[i]

    current_row = start_row + 1
    for _, row_data in df_summary.iterrows():
        for i, col_name in enumerate(headers):
            value = row_data[col_name]
            cell = worksheet.cell(row=current_row, column=start_col + i)
            
            if "R$" in col_name and isinstance(value, (int, float)):
                cell.value = value
                cell.number_format = 'R$ #,##0.00'
                cell.alignment = RIGHT_ALIGN
            elif "%" in col_name and isinstance(value, (int, float)):
                cell.value = value / 100
                cell.number_format = '0.00"%"'
                cell.alignment = CENTER_ALIGN
            else:
                cell.value = value
                if "Qtd." in col_name or "Qtd_" in col_name:
                    cell.alignment = CENTER_ALIGN
                else:
                    cell.alignment = LEFT_ALIGN

            cell.border = THIN_BORDER
            if str(row_data.get(headers[0])).upper() == 'TOTAL':
                cell.font = Font(bold=True)
                cell.fill = KPI_FILL
        current_row += 1
    return current_row

def _create_situacao_summary(df):
    summary = df.groupby('Situação').agg(
        Qtd_Notas=('Número da Nota', 'count'),
        Valor_Servico=('Valor do Serviço (R$)', 'sum')
    ).reset_index().rename(columns={'Valor_Servico': 'Valor Serviço (R$)'})
    total_row = summary.sum(numeric_only=True)
    total_row['Situação'] = 'TOTAL'
    summary = pd.concat([summary, pd.DataFrame([total_row])], ignore_index=True)
    headers = ['Situação', 'Qtd_Notas', 'Valor Serviço (R$)']
    widths = [18, 12, 20]
    return summary, headers, widths

def _create_codigo_servico_summary(df):
    df_validas = df[df['Situação'] != 'CANCELADA'].copy()
    summary = df_validas.groupby('Código do Serviço').agg(
        Qtd_Notas=('Número da Nota', 'count'),
        Valor_Servico=('Valor do Serviço (R$)', 'sum'),
        ISS_Proprio=('ISS Próprio (R$)', 'sum'),
        ISS_Retido=('ISS Retido (R$)', 'sum'),
        Aliquota_perc=('Alíquota (%)', 'first')
    ).reset_index().rename(columns={
        'Valor_Servico': 'Valor Serviço (R$)', 'ISS_Proprio': 'ISS Próprio (R$)',
        'ISS_Retido': 'ISS Retido (R$)', 'Aliquota_perc': 'Alíquota (%)'
    })
    
    # Calcular totais (excluindo alíquota que não deve ser somada)
    numeric_cols = summary.select_dtypes(include=['number']).columns.tolist()
    if 'Alíquota (%)' in numeric_cols:
        numeric_cols.remove('Alíquota (%)')
    total_row = summary[numeric_cols].sum()
    total_row['Código do Serviço'] = 'TOTAL'
    
    # Não calcular alíquota para o total - deixar vazio
    total_row['Alíquota (%)'] = 'N/A'
    
    # Formatar alíquotas como percentual após os cálculos
    summary['Alíquota (%)'] = summary['Alíquota (%)'].apply(lambda x: f"{x:.2f}%" if pd.notna(x) else "0.00%")
    
    summary = pd.concat([summary, pd.DataFrame([total_row])], ignore_index=True)
    headers = ['Código do Serviço', 'Qtd_Notas', 'Valor Serviço (R$)', 'ISS Próprio (R$)', 'ISS Retido (R$)', 'Alíquota (%)']
    widths = [18, 12, 20, 18, 18, 12]
    return summary, headers, widths

def _create_top_tomadores_summary(df):
    df_validas = df[df['Situação'] != 'CANCELADA'].copy()
    summary = df_validas.groupby('Nome do Tomador').agg(
        Valor_Total=('Valor do Serviço (R$)', 'sum'),
        Qtd_Notas=('Número da Nota', 'count')
    ).reset_index()
    summary = summary.sort_values('Valor_Total', ascending=False).head(5).rename(columns={'Valor_Total': 'Valor Total (R$)'})
    headers = ['Nome do Tomador', 'Qtd_Notas', 'Valor Total (R$)']
    widths = [40, 12, 20]
    return summary, headers, widths

def _create_situacao_pie_chart(worksheet, summary, anchor_cell, title, data_start_row, data_start_col):
    pie = PieChart()
    pie.title = title
    pie.style = 26
    
    labels_data = summary[summary['Situação'] != 'TOTAL']
    
    # Referências corrigidas
    labels = Reference(worksheet, min_col=data_start_col, min_row=data_start_row + 1, max_row=data_start_row + len(labels_data))
    data = Reference(worksheet, min_col=data_start_col + 2, min_row=data_start_row, max_row=data_start_row + len(labels_data))
    
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    
    # Adicionar percentuais no gráfico de pizza
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.dataLabels.showVal = False
    pie.dataLabels.showCatName = True
    
    worksheet.add_chart(pie, anchor_cell)

def _create_top_tomadores_bar_chart(worksheet, summary, anchor_cell, title, data_start_row, data_start_col):
    chart = BarChart()
    chart.title = title
    chart.style = 4
    chart.y_axis.title = 'Valor Total (R$)'
    chart.x_axis.title = 'Tomador'

    # Referências corrigidas
    data = Reference(worksheet, min_col=data_start_col + 2, min_row=data_start_row, max_row=data_start_row + len(summary))
    cats = Reference(worksheet, min_col=data_start_col, min_row=data_start_row + 1, max_row=data_start_row + len(summary))
    
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.legend = None
    
    # Adicionar rótulos de dados com nomes dos tomadores
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = True
    chart.dataLabels.showCatName = False

    worksheet.add_chart(chart, anchor_cell)

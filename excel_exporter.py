import pandas as pd
from datetime import datetime
import logging
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("excel_exporter.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("excel_exporter")

def format_codigo_servico(codigo):
    """
    Mantém o código do serviço no formato original do PDF.

    Args:
        codigo: O código do serviço a ser mantido no formato original

    Returns:
        str: O código no formato original
    """
    # Simplesmente retornar o código original
    return codigo

def export_to_excel(df, output_path=None):
    """
    Exporta os dados para um arquivo Excel.

    Args:
        df (pd.DataFrame): DataFrame com os dados a serem exportados
        output_path (str, optional): Caminho para salvar o arquivo Excel

    Returns:
        str: Caminho do arquivo Excel gerado
    """
    if df.empty:
        logger.warning("DataFrame vazio. Nenhum arquivo Excel gerado.")
        return None

    # Verificar e corrigir o campo Incidência
    if 'Incidência' in df.columns:
        # Verificar se o valor está truncado
        if (df['Incidência'] == 'ESTAB. DO').any():
            logger.info("Corrigindo campo Incidência...")
            df['Incidência'] = 'ESTAB. DO PRESTADOR'

    # Criar uma cópia do DataFrame para não modificar o original
    df_excel = df.copy()

    # Verificar linhas sem tomador
    if 'Nome do Tomador' in df_excel.columns:
        # Adicionar coluna de verificação para linhas sem tomador
        df_excel['Verificação'] = 'OK'
        # Marcar linhas sem tomador
        sem_tomador = df_excel['Nome do Tomador'].isna() | (df_excel['Nome do Tomador'] == '')
        df_excel.loc[sem_tomador, 'Verificação'] = 'SEM TOMADOR'

        # Contar e logar o número de linhas sem tomador
        num_sem_tomador = sem_tomador.sum()
        logger.warning(f"Encontradas {num_sem_tomador} linhas sem tomador. Estas linhas foram marcadas na coluna 'Verificação'.")

    # Formatar o código do serviço na aba Dados_NFSe
    if 'Código do Serviço' in df_excel.columns:
        # Aplicar a formatação apenas para valores que não são TOTAL
        df_excel['Código do Serviço Formatado'] = df_excel['Código do Serviço'].apply(
            lambda x: format_codigo_servico(x) if x != 'TOTAL' else x
        )
        # Substituir a coluna original pela formatada
        df_excel['Código do Serviço'] = df_excel['Código do Serviço Formatado']
        # Remover a coluna temporária
        df_excel.drop('Código do Serviço Formatado', axis=1, inplace=True)

    # Garantir que a alíquota seja tratada corretamente
    if 'Alíquota (%)' in df_excel.columns:
        # Converter para numérico e garantir que seja um valor percentual
        df_excel['Alíquota (%)'] = pd.to_numeric(df_excel['Alíquota (%)'], errors='coerce').fillna(0.0)

        # Verificar se há valores muito altos (ex: 300 para 3%)
        mask = df_excel['Alíquota (%)'] > 100
        df_excel.loc[mask, 'Alíquota (%)'] = df_excel.loc[mask, 'Alíquota (%)'] / 100

    # Se o caminho não for fornecido, criar um com timestamp
    if output_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = f"notas_fiscais_{timestamp}.xlsx"

    logger.info(f"Iniciando exportação para Excel: {output_path}")

    try:
        # Criar um escritor Excel com o engine openpyxl
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Criar o workbook se não existir
            if not hasattr(writer, 'book'):
                writer.book = writer.book

            # Planilha com todos os dados - renomeada para "Dados_NFSe"
            df_excel.to_excel(writer, sheet_name='Dados_NFSe', index=False)
            format_excel_sheet(writer, 'Dados_NFSe', df_excel)

            # Adicionar planilha de sumário - renomeada para "Resumo_NFSe"
            create_summary_sheet(writer, df_excel, sheet_name='Resumo_NFSe')

        logger.info(f"Exportação para Excel concluída: {output_path}")
        return output_path

    except Exception as e:
        logger.error(f"Erro ao exportar para Excel: {str(e)}")
        raise

def format_excel_sheet(writer, sheet_name, df):
    """
    Formata uma planilha Excel com estilos aprimorados.

    Args:
        writer (pd.ExcelWriter): Escritor Excel
        sheet_name (str): Nome da planilha
        df (pd.DataFrame): DataFrame com os dados
    """
    # Obter a planilha
    worksheet = writer.sheets[sheet_name]

    # Definir largura das colunas
    for i, col in enumerate(df.columns):
        # Calcular a largura máxima da coluna
        max_len = max(
            df[col].astype(str).map(len).max(),  # Comprimento máximo dos dados
            len(str(col))  # Comprimento do cabeçalho
        ) + 2  # Adicionar um pouco de espaço extra

        # Definir a largura da coluna (converter para unidades do Excel)
        column_letter = get_column_letter(i + 1)  # Usando a função importada para obter a letra da coluna
        worksheet.column_dimensions[column_letter].width = max_len

    # Formatar cabeçalho com estilo aprimorado
    header_row = worksheet[worksheet.min_row]
    for cell in header_row:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Identificar colunas por tipo para formatação
    # Colunas monetárias
    monetary_columns = [
        'Valor do Serviço (R$)',
        'Base de Cálculo (R$)',
        'ISS Próprio (R$)',
        'ISS Retido (R$)',
        'Valor Médio por Nota (R$)'
    ]

    # Colunas percentuais
    percentage_columns = [
        'Alíquota (%)',
        'Proporção ISS Retido (%)',
        'Crescimento Valor (%)'
    ]

    # Mapear nomes de colunas para índices
    column_indices = {col: i for i, col in enumerate(df.columns)}

    # Aplicar formatação para cada linha de dados
    for row_idx in range(worksheet.min_row + 1, worksheet.max_row + 1):
        # Verificar se a linha atual tem a coluna de verificação e se está marcada como "SEM TOMADOR"
        verificacao_idx = column_indices.get('Verificação', None)
        is_sem_tomador = False

        if verificacao_idx is not None:
            verificacao_cell = worksheet[f"{get_column_letter(verificacao_idx + 1)}{row_idx}"]
            if verificacao_cell.value == 'SEM TOMADOR':
                is_sem_tomador = True

        # Definir cor de fundo com base na verificação e alternância de linhas
        if is_sem_tomador:
            # Destacar linhas sem tomador com cor de fundo amarela
            row_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        else:
            # Alternar cores de fundo para melhorar legibilidade
            row_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid") if row_idx % 2 == 0 else None

        for col_name, col_idx in column_indices.items():
            col_letter = get_column_letter(col_idx + 1)
            cell = worksheet[f"{col_letter}{row_idx}"]

            # Aplicar preenchimento
            if row_fill:
                cell.fill = row_fill

            # Adicionar comentário para linhas sem tomador
            if is_sem_tomador and col_name == 'Nome do Tomador':
                from openpyxl.comments import Comment
                cell.comment = Comment("Esta linha não possui registro de tomador e pode conter dados incorretos.", "Sistema")

            # Centralizar texto em algumas colunas
            if col_name in ['Competência', 'Quantidade de Notas']:
                cell.alignment = Alignment(horizontal='center')

            # Formatar valores monetários usando o formato de moeda simples
            if col_name in monetary_columns:
                # Usar formato de moeda simples (código 44 no Excel)
                cell.number_format = 'R$ #,##0.00'
                cell.alignment = Alignment(horizontal='right')

            # Formatar colunas percentuais
            elif col_name in percentage_columns:
                # Garantir que o valor seja um número e esteja no formato correto
                if cell.value is not None:
                    try:
                        # Converter para float
                        value = float(cell.value)

                        # Se o valor for muito alto (ex: 300 para 3%), dividir por 100
                        if value > 100:
                            value = value / 100

                        # Definir o valor da célula
                        cell.value = value

                        # Usar formato numérico com 2 casas decimais seguido de %
                        cell.number_format = '0.00"%"'
                    except (ValueError, TypeError):
                        # Se não for possível converter, manter o valor original
                        pass

                # Centralizar o valor
                cell.alignment = Alignment(horizontal='center')

                # Formatação especial para crescimento (verde positivo, vermelho negativo)
                if col_name == 'Crescimento Valor (%)' and cell.value is not None:
                    try:
                        value = float(cell.value)
                        if value > 0:
                            cell.font = Font(color="006100")  # Verde escuro
                        elif value < 0:
                            cell.font = Font(color="9C0006")  # Vermelho escuro
                    except (ValueError, TypeError):
                        pass

    # Adicionar bordas externas à tabela
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Aplicar bordas a todas as células com dados
    for row in worksheet.iter_rows(min_row=worksheet.min_row, max_row=worksheet.max_row):
        for cell in row:
            if cell.value is not None:  # Aplicar apenas a células com conteúdo
                cell.border = thin_border

def create_summary_sheet(writer, df, sheet_name='Resumo_NFSe'):
    """
    Cria uma planilha de resumo com tabelas organizadas por assunto, sem gráficos.

    Args:
        writer (pd.ExcelWriter): Escritor Excel
        df (pd.DataFrame): DataFrame com os dados
        sheet_name (str, optional): Nome da planilha de resumo. Padrão é 'Resumo_NFSe'.
    """
    logger.info(f"Criando planilha de resumo organizada: {sheet_name}")

    # Verificar se há dados suficientes
    if df.empty:
        logger.warning("DataFrame vazio. Nenhum resumo gerado.")
        return

    # Garantir que temos as colunas necessárias
    required_columns = ['Número da Nota', 'Valor do Serviço (R$)', 'Base de Cálculo (R$)',
                        'ISS Próprio (R$)', 'ISS Retido (R$)', 'Situação']

    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        logger.warning(f"Colunas necessárias ausentes: {missing_columns}")
        return

    # Garantir que os valores numéricos estão no formato correto
    numeric_columns = ['Valor do Serviço (R$)', 'Base de Cálculo (R$)',
                      'ISS Próprio (R$)', 'ISS Retido (R$)', 'Alíquota (%)']

    for col in numeric_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

            # Arredondar para 2 casas decimais
            if col != 'Alíquota (%)':
                df[col] = df[col].round(2)

    # Criar planilha
    worksheet = writer.book.create_sheet(sheet_name)

    # Adicionar título principal
    worksheet.merge_cells('A1:J1')
    title_cell = worksheet['A1']
    title_cell.value = "RESUMO DE NOTAS FISCAIS DE SERVIÇO"
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    title_cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    title_cell.font = Font(bold=True, size=14, color="FFFFFF")

    # Definir altura da linha do título
    worksheet.row_dimensions[1].height = 25

    # Adicionar data de geração do relatório
    worksheet['J2'] = f"Relatório: {datetime.now().strftime('%d/%m/%Y')}"
    worksheet['J2'].font = Font(italic=True, size=8)
    worksheet['J2'].alignment = Alignment(horizontal='right')

    # Definir linha atual para adicionar as tabelas
    current_row = 4  # Começar na linha 4 após o título e data

    # ===== TABELA 1: RESUMO POR SITUAÇÃO =====
    # Adicionar título da seção
    worksheet.merge_cells(f'A{current_row}:F{current_row}')
    section_title = worksheet[f'A{current_row}']
    section_title.value = "RESUMO POR SITUAÇÃO"
    section_title.font = Font(bold=True, size=12)
    section_title.alignment = Alignment(horizontal='center', vertical='center')
    section_title.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    section_title.font = Font(bold=True, size=12, color="FFFFFF")

    # Definir altura da linha do título da seção
    worksheet.row_dimensions[current_row].height = 20

    current_row += 1

    # Calcular resumo por situação
    situacao_summary = df.groupby('Situação').agg({
        'Número da Nota': 'count',
        'Valor do Serviço (R$)': 'sum',
        'Base de Cálculo (R$)': 'sum',
        'ISS Próprio (R$)': 'sum',
        'ISS Retido (R$)': 'sum'
    }).reset_index()

    # Renomear a coluna de contagem
    situacao_summary = situacao_summary.rename(columns={'Número da Nota': 'Quantidade de Notas'})

    # Arredondar valores para 2 casas decimais
    for col in ['Valor do Serviço (R$)', 'Base de Cálculo (R$)', 'ISS Próprio (R$)', 'ISS Retido (R$)']:
        situacao_summary[col] = situacao_summary[col].round(2)

    # Adicionar linha de VALIDADAS (soma de ESCRITURADA e QUITADA)
    validadas_situacao = {
        'Situação': 'VALIDADAS',
        'Quantidade de Notas': 0,
        'Valor do Serviço (R$)': 0,
        'Base de Cálculo (R$)': 0,
        'ISS Próprio (R$)': 0,
        'ISS Retido (R$)': 0
    }

    # Filtrar apenas as linhas ESCRITURADA e QUITADA
    validadas_df = situacao_summary[situacao_summary['Situação'].isin(['ESCRITURADA', 'QUITADA'])]

    if not validadas_df.empty:
        validadas_situacao['Quantidade de Notas'] = validadas_df['Quantidade de Notas'].sum()
        validadas_situacao['Valor do Serviço (R$)'] = validadas_df['Valor do Serviço (R$)'].sum().round(2)
        validadas_situacao['Base de Cálculo (R$)'] = validadas_df['Base de Cálculo (R$)'].sum().round(2)
        validadas_situacao['ISS Próprio (R$)'] = validadas_df['ISS Próprio (R$)'].sum().round(2)
        validadas_situacao['ISS Retido (R$)'] = validadas_df['ISS Retido (R$)'].sum().round(2)

    # Adicionar linha de total
    total_situacao = {
        'Situação': 'TOTAL',
        'Quantidade de Notas': situacao_summary['Quantidade de Notas'].sum(),
        'Valor do Serviço (R$)': situacao_summary['Valor do Serviço (R$)'].sum().round(2),
        'Base de Cálculo (R$)': situacao_summary['Base de Cálculo (R$)'].sum().round(2),
        'ISS Próprio (R$)': situacao_summary['ISS Próprio (R$)'].sum().round(2),
        'ISS Retido (R$)': situacao_summary['ISS Retido (R$)'].sum().round(2)
    }

    # Adicionar as linhas de VALIDADAS e TOTAL ao DataFrame
    situacao_summary = pd.concat([situacao_summary, pd.DataFrame([validadas_situacao, total_situacao])])

    # Adicionar cabeçalhos
    headers = ['Situação', 'Quantidade de Notas', 'Valor do Serviço (R$)',
               'Base de Cálculo (R$)', 'ISS Próprio (R$)', 'ISS Retido (R$)']

    # Definir larguras das colunas para a tabela de situação
    column_widths = {
        'A': 15,  # Situação
        'B': 15,  # Quantidade de Notas
        'C': 20,  # Valor do Serviço
        'D': 20,  # Base de Cálculo
        'E': 15,  # ISS Próprio
        'F': 15,  # ISS Retido
    }

    for col_letter, width in column_widths.items():
        worksheet.column_dimensions[col_letter].width = width

    for i, header in enumerate(headers):
        cell = worksheet.cell(row=current_row, column=i+1)
        cell.value = header
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        cell.alignment = Alignment(horizontal='center')
        cell.border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )

    current_row += 1

    # Adicionar dados
    for _, row in situacao_summary.iterrows():
        for i, col in enumerate(headers):
            cell = worksheet.cell(row=current_row, column=i+1)

            # Formatação especial para valores monetários
            if 'R$' in col:
                # Armazenar o valor como número, não como string, e arredondar para 2 casas decimais
                try:
                    cell.value = round(float(row[col]), 2)
                except (ValueError, TypeError):
                    # Se não for possível converter para float, usar o valor original
                    if isinstance(row[col], (int, float)):
                        cell.value = row[col]
                    else:
                        # Tentar converter para float novamente
                        try:
                            cell.value = float(str(row[col]).replace('R$', '').replace('.', '').replace(',', '.').strip())
                        except:
                            cell.value = 0.0
                # Usar formato de moeda simples (código 44 no Excel)
                cell.number_format = 'R$ #,##0.00'
                cell.alignment = Alignment(horizontal='right')
            elif col == 'Alíquota (%)':
                # Verificar se o valor não é None
                if pd.notna(row[col]):
                    # Obter o valor e garantir que seja um número
                    value = float(row[col])
                    # Se o valor for muito alto (ex: 300), dividir por 100 para obter 3
                    if value > 100:
                        value = value / 100
                    cell.value = value
                    # Usar formato numérico com 2 casas decimais seguido de %
                    cell.number_format = '0.00"%"'
                cell.alignment = Alignment(horizontal='center')
            elif col == 'Quantidade de Notas':
                cell.value = int(row[col])
                cell.alignment = Alignment(horizontal='center')
            else:
                cell.value = row[col]

            # Formatação especial para as linhas de VALIDADAS e TOTAL
            if row['Situação'] == 'TOTAL':
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
            elif row['Situação'] == 'VALIDADAS':
                cell.font = Font(bold=True, italic=True)
                cell.fill = PatternFill(start_color="E6F0FF", end_color="E6F0FF", fill_type="solid")

            # Adicionar bordas
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )

        current_row += 1

    # Espaço entre tabelas
    current_row += 2

    # ===== TABELA 2: RESUMO POR COMPETÊNCIA =====
    # Verificar se temos as colunas de competência
    if 'Competência' in df.columns or ('Mês' in df.columns and 'Ano' in df.columns):
        # Adicionar título da seção
        worksheet.merge_cells(f'A{current_row}:I{current_row}')
        section_title = worksheet[f'A{current_row}']
        section_title.value = "RESUMO POR COMPETÊNCIA (VALIDADAS, CANCELADAS E TOTAL)"
        section_title.font = Font(bold=True, size=12)
        section_title.alignment = Alignment(horizontal='center', vertical='center')
        section_title.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        section_title.font = Font(bold=True, size=12, color="FFFFFF")

        # Definir altura da linha do título da seção
        worksheet.row_dimensions[current_row].height = 20

        current_row += 1

        # Preparar dados de competência
        if 'Competência' not in df.columns and 'Mês' in df.columns and 'Ano' in df.columns:
            # Criar coluna de competência a partir de Mês e Ano
            df['Competência'] = df.apply(
                lambda x: f"{int(x['Mês']):02d}/{int(x['Ano'])}"
                if pd.notna(x['Mês']) and pd.notna(x['Ano'])
                else "Desconhecida",
                axis=1
            )

        # Separar notas validadas e canceladas
        df_validadas = df[df['Situação'] != 'CANCELADA'].copy()
        df_canceladas = df[df['Situação'] == 'CANCELADA'].copy()
        
        # Obter todas as competências únicas do DataFrame original
        todas_competencias = df['Competência'].unique()
        
        # Calcular resumo por competência para notas validadas
        validadas_summary = df_validadas.groupby('Competência').agg({
            'Número da Nota': 'count',
            'Valor do Serviço (R$)': 'sum',
            'Base de Cálculo (R$)': 'sum',
            'ISS Próprio (R$)': 'sum',
            'ISS Retido (R$)': 'sum'
        }).reset_index()
        validadas_summary['Tipo'] = 'VALIDADAS'
        validadas_summary = validadas_summary.rename(columns={'Número da Nota': 'Quantidade de Notas'})
        
        # Adicionar competências faltantes para validadas (com valores zerados)
        competencias_faltantes_validadas = set(todas_competencias) - set(validadas_summary['Competência'])
        for comp in competencias_faltantes_validadas:
            linha_zerada = {
                'Competência': comp,
                'Tipo': 'VALIDADAS',
                'Quantidade de Notas': 0,
                'Valor do Serviço (R$)': 0.0,
                'Base de Cálculo (R$)': 0.0,
                'ISS Próprio (R$)': 0.0,
                'ISS Retido (R$)': 0.0
            }
            validadas_summary = pd.concat([validadas_summary, pd.DataFrame([linha_zerada])])
        
        # Calcular resumo por competência para notas canceladas
        canceladas_summary = df_canceladas.groupby('Competência').agg({
            'Número da Nota': 'count',
            'Valor do Serviço (R$)': 'sum',
            'Base de Cálculo (R$)': 'sum',
            'ISS Próprio (R$)': 'sum',
            'ISS Retido (R$)': 'sum'
        }).reset_index()
        canceladas_summary['Tipo'] = 'CANCELADAS'
        canceladas_summary = canceladas_summary.rename(columns={'Número da Nota': 'Quantidade de Notas'})
        
        # Adicionar competências faltantes para canceladas (com valores zerados)
        competencias_faltantes_canceladas = set(todas_competencias) - set(canceladas_summary['Competência'])
        for comp in competencias_faltantes_canceladas:
            linha_zerada = {
                'Competência': comp,
                'Tipo': 'CANCELADAS',
                'Quantidade de Notas': 0,
                'Valor do Serviço (R$)': 0.0,
                'Base de Cálculo (R$)': 0.0,
                'ISS Próprio (R$)': 0.0,
                'ISS Retido (R$)': 0.0
            }
            canceladas_summary = pd.concat([canceladas_summary, pd.DataFrame([linha_zerada])])
        
        # Calcular resumo por competência para todas as notas (total)
        total_summary = df.groupby('Competência').agg({
            'Número da Nota': 'count',
            'Valor do Serviço (R$)': 'sum',
            'Base de Cálculo (R$)': 'sum',
            'ISS Próprio (R$)': 'sum',
            'ISS Retido (R$)': 'sum'
        }).reset_index()
        total_summary['Tipo'] = 'TOTAL'
        total_summary = total_summary.rename(columns={'Número da Nota': 'Quantidade de Notas'})
        
        # Combinar os três DataFrames
        competencia_summary = pd.concat([validadas_summary, canceladas_summary, total_summary])

        # Arredondar valores para 2 casas decimais
        for col in ['Valor do Serviço (R$)', 'Base de Cálculo (R$)', 'ISS Próprio (R$)', 'ISS Retido (R$)']:
            competencia_summary[col] = competencia_summary[col].round(2)

        # Adicionar coluna de valor médio por nota (evitando divisão por zero)
        competencia_summary['Valor Médio por Nota (R$)'] = (
            competencia_summary['Valor do Serviço (R$)'] / competencia_summary['Quantidade de Notas'].replace(0, float('nan'))
        ).round(2)

        # Adicionar coluna de ISS Total (soma do ISS Próprio e ISS Retido)
        competencia_summary['ISS Total (R$)'] = (
            competencia_summary['ISS Próprio (R$)'] + competencia_summary['ISS Retido (R$)']
        ).round(2)

        # Total ISS removido conforme solicitado

        # Adicionar coluna de ano para agrupar por ano
        try:
            # Tentar extrair o ano da competência (formato MM/AAAA)
            competencia_summary['temp_date'] = pd.to_datetime(
                competencia_summary['Competência'],
                format='%m/%Y',
                errors='coerce'
            )
            competencia_summary['Ano'] = competencia_summary['temp_date'].dt.year

            # Ordenar por data e tipo (VALIDADAS, CANCELADAS, TOTAL)
            competencia_summary = competencia_summary.sort_values(['temp_date', 'Tipo'])

            # Criar um DataFrame para armazenar as linhas finais (incluindo subtotais e total geral)
            final_df = []

            # Agrupar por ano e calcular subtotais
            anos = competencia_summary['Ano'].dropna().unique()

            for ano in anos:
                # Filtrar as competências do ano atual
                df_ano = competencia_summary[competencia_summary['Ano'] == ano].copy()
                
                # Reorganizar para que cada competência tenha suas três linhas juntas (VALIDADAS, CANCELADAS, TOTAL)
                competencias_ano = df_ano['Competência'].drop_duplicates().tolist()
                df_ano_reorganizado = []
                
                for comp in competencias_ano:
                    # Filtrar apenas as linhas da competência atual
                    df_comp = df_ano[df_ano['Competência'] == comp].copy()
                    
                    # Verificar quais tipos existem para esta competência
                    tipos_existentes = df_comp['Tipo'].unique()
                    
                    # Garantir que os três tipos estejam presentes, na ordem correta
                    df_comp_organizado = []
                    
                    for tipo in ['VALIDADAS', 'CANCELADAS', 'TOTAL']:
                        if tipo in tipos_existentes:
                            # Se o tipo existe, adicionar as linhas correspondentes
                            df_comp_organizado.append(df_comp[df_comp['Tipo'] == tipo])
                        else:
                            # Se o tipo não existe, criar uma linha com valores zerados
                            linha_zerada = {
                                'Competência': comp,
                                'Tipo': tipo,
                                'Quantidade de Notas': 0,
                                'Valor do Serviço (R$)': 0.0,
                                'Valor Médio por Nota (R$)': 0.0,
                                'Base de Cálculo (R$)': 0.0,
                                'ISS Próprio (R$)': 0.0,
                                'ISS Retido (R$)': 0.0,
                                'ISS Total (R$)': 0.0,
                                'Ano': df_comp['Ano'].iloc[0] if not df_comp.empty else ano
                            }
                            df_comp_organizado.append(pd.DataFrame([linha_zerada]))
                    
                    # Concatenar os três tipos para esta competência
                    if df_comp_organizado:
                        df_comp_final = pd.concat(df_comp_organizado)
                        df_ano_reorganizado.append(df_comp_final)
                
                # Concatenar todas as competências reorganizadas
                if df_ano_reorganizado:
                    df_ano_final = pd.concat(df_ano_reorganizado)
                    final_df.append(df_ano_final)
                
                # Calcular subtotais do ano para cada tipo
                for tipo in ['VALIDADAS', 'CANCELADAS', 'TOTAL']:
                    df_ano_tipo = df_ano[df_ano['Tipo'] == tipo].copy()
                    
                    if not df_ano_tipo.empty:
                        subtotal_ano = {
                            'Competência': f'SUBTOTAL {int(ano)}',
                            'Quantidade de Notas': df_ano_tipo['Quantidade de Notas'].sum(),
                            'Valor do Serviço (R$)': df_ano_tipo['Valor do Serviço (R$)'].sum().round(2),
                            'Valor Médio por Nota (R$)': (
                                df_ano_tipo['Valor do Serviço (R$)'].sum() /
                                df_ano_tipo['Quantidade de Notas'].sum()
                            ).round(2) if df_ano_tipo['Quantidade de Notas'].sum() > 0 else 0,
                            'Base de Cálculo (R$)': df_ano_tipo['Base de Cálculo (R$)'].sum().round(2),
                            'ISS Próprio (R$)': df_ano_tipo['ISS Próprio (R$)'].sum().round(2),
                            'ISS Retido (R$)': df_ano_tipo['ISS Retido (R$)'].sum().round(2),
                            'ISS Total (R$)': df_ano_tipo['ISS Total (R$)'].sum().round(2),
                            'Tipo': tipo,
                            'Ano': ano  # Manter o ano para referência
                        }
                        
                        # Adicionar o subtotal ao DataFrame final
                        final_df.append(pd.DataFrame([subtotal_ano]))

            # Concatenar todos os DataFrames
            competencia_summary = pd.concat(final_df)

            # Remover a coluna temporária de data e ano
            competencia_summary = competencia_summary.drop(['temp_date', 'Ano'], axis=1)
        except Exception as e:
            # Se falhar, manter a ordem original e não adicionar subtotais
            logger.warning(f"Não foi possível adicionar subtotais por ano: {e}")
            if 'temp_date' in competencia_summary.columns:
                competencia_summary = competencia_summary.drop('temp_date', axis=1)
            if 'Ano' in competencia_summary.columns:
                competencia_summary = competencia_summary.drop('Ano', axis=1)

        # Adicionar linhas de total geral para cada tipo
        for tipo in ['VALIDADAS', 'CANCELADAS', 'TOTAL']:
            # Filtrar apenas as linhas do tipo atual e que não são subtotais
            df_tipo = competencia_summary[
                (competencia_summary['Tipo'] == tipo) & 
                (~competencia_summary['Competência'].str.contains('SUBTOTAL', na=False))
            ]
            
            if not df_tipo.empty:
                total_competencia = {
                    'Competência': 'TOTAL GERAL',
                    'Quantidade de Notas': df_tipo['Quantidade de Notas'].sum(),
                    'Valor do Serviço (R$)': df_tipo['Valor do Serviço (R$)'].sum().round(2),
                    'Valor Médio por Nota (R$)': (
                        df_tipo['Valor do Serviço (R$)'].sum() /
                        df_tipo['Quantidade de Notas'].sum()
                    ).round(2) if df_tipo['Quantidade de Notas'].sum() > 0 else 0,
                    'Base de Cálculo (R$)': df_tipo['Base de Cálculo (R$)'].sum().round(2),
                    'ISS Próprio (R$)': df_tipo['ISS Próprio (R$)'].sum().round(2),
                    'ISS Retido (R$)': df_tipo['ISS Retido (R$)'].sum().round(2),
                    'ISS Total (R$)': df_tipo['ISS Total (R$)'].sum().round(2),
                    'Tipo': tipo
                }
                
                # Adicionar o total geral ao DataFrame
                competencia_summary = pd.concat([competencia_summary, pd.DataFrame([total_competencia])])

        # Adicionar cabeçalhos
        headers = [
            'Competência', 'Tipo', 'Quantidade de Notas', 'Valor do Serviço (R$)',
            'Valor Médio por Nota (R$)', 'Base de Cálculo (R$)',
            'ISS Próprio (R$)', 'ISS Retido (R$)', 'ISS Total (R$)'
        ]

        # Definir larguras das colunas para a tabela de competência
        column_widths = {
            'A': 15,  # Competência
            'B': 12,  # Tipo
            'C': 15,  # Quantidade de Notas
            'D': 20,  # Valor do Serviço
            'E': 20,  # Valor Médio por Nota
            'F': 20,  # Base de Cálculo
            'G': 15,  # ISS Próprio
            'H': 15,  # ISS Retido
            'I': 15,  # ISS Total
        }

        for col_letter, width in column_widths.items():
            worksheet.column_dimensions[col_letter].width = width

        for i, header in enumerate(headers):
            cell = worksheet.cell(row=current_row, column=i+1)
            cell.value = header
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
            cell.alignment = Alignment(horizontal='center')
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )

        current_row += 1

        # Adicionar dados
        for _, row in competencia_summary.iterrows():
            for i, col in enumerate(headers):
                cell = worksheet.cell(row=current_row, column=i+1)

                # Formatação especial para valores monetários
                if 'R$' in col:
                    # Armazenar o valor como número, não como string, e arredondar para 2 casas decimais
                    try:
                        cell.value = round(float(row[col]), 2)
                    except (ValueError, TypeError):
                        # Se não for possível converter para float, usar o valor original
                        if isinstance(row[col], (int, float)):
                            cell.value = row[col]
                        else:
                            # Tentar converter para float novamente
                            try:
                                cell.value = float(str(row[col]).replace('R$', '').replace('.', '').replace(',', '.').strip())
                            except:
                                cell.value = 0.0
                    # Usar formato de moeda simples (código 44 no Excel)
                    cell.number_format = 'R$ #,##0.00'
                    cell.alignment = Alignment(horizontal='right')
                elif col == 'Alíquota (%)':
                    # Verificar se o valor não é None
                    if pd.notna(row[col]):
                        # Obter o valor e garantir que seja um número
                        value = float(row[col])
                        # Se o valor for muito alto (ex: 300), dividir por 100 para obter 3
                        if value > 100:
                            value = value / 100
                        cell.value = value
                        # Usar formato numérico com 2 casas decimais seguido de %
                        cell.number_format = '0.00"%"'
                    cell.alignment = Alignment(horizontal='center')
                elif col == 'Quantidade de Notas':
                    cell.value = int(row[col])
                    cell.alignment = Alignment(horizontal='center')
                elif col == 'Competência':
                    cell.value = row[col]
                    cell.alignment = Alignment(horizontal='center')
                else:
                    cell.value = row[col]

                # Formatação especial para as linhas de subtotal e total
                if row['Competência'] == 'TOTAL GERAL':
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
                elif 'SUBTOTAL' in str(row['Competência']):
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="EBF1DE", end_color="EBF1DE", fill_type="solid")
                
                # Formatação especial para os tipos de notas
                if col == 'Tipo':
                    cell.alignment = Alignment(horizontal='center')
                    if row['Tipo'] == 'VALIDADAS':
                        cell.fill = PatternFill(start_color="D8E4BC", end_color="D8E4BC", fill_type="solid")
                    elif row['Tipo'] == 'CANCELADAS':
                        cell.fill = PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid")
                    elif row['Tipo'] == 'TOTAL':
                        cell.fill = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
                        cell.font = Font(bold=True)

                # Adicionar bordas
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )

            current_row += 1

    # Espaço entre tabelas
    current_row += 2

    # ===== TABELA 3: RESUMO POR CÓDIGO DO SERVIÇO =====
    # Verificar se temos a coluna de código do serviço
    if 'Código do Serviço' in df.columns:
        # Adicionar título da seção
        worksheet.merge_cells(f'A{current_row}:G{current_row}')
        section_title = worksheet[f'A{current_row}']
        section_title.value = "RESUMO POR CÓDIGO DO SERVIÇO"
        section_title.font = Font(bold=True, size=12)
        section_title.alignment = Alignment(horizontal='center', vertical='center')
        section_title.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        section_title.font = Font(bold=True, size=12, color="FFFFFF")

        # Definir altura da linha do título da seção
        worksheet.row_dimensions[current_row].height = 20

        current_row += 1

        # Filtrar notas canceladas antes de calcular o resumo por código do serviço
        df_sem_canceladas = df[df['Situação'] != 'CANCELADA'].copy()

        # Calcular resumo por código do serviço (excluindo notas canceladas)
        codigo_servico_summary = df_sem_canceladas.groupby('Código do Serviço').agg({
            'Número da Nota': 'count',
            'Valor do Serviço (R$)': 'sum',
            'Base de Cálculo (R$)': 'sum',
            'ISS Próprio (R$)': 'sum',
            'ISS Retido (R$)': 'sum',
            'Alíquota (%)': 'first'  # Pegar a primeira alíquota para cada código
        }).reset_index()

        # Renomear a coluna de contagem
        codigo_servico_summary = codigo_servico_summary.rename(columns={'Número da Nota': 'Quantidade de Notas'})

        # Arredondar valores para 2 casas decimais
        for col in ['Valor do Serviço (R$)', 'Base de Cálculo (R$)', 'ISS Próprio (R$)', 'ISS Retido (R$)']:
            codigo_servico_summary[col] = codigo_servico_summary[col].round(2)

        # Adicionar linha de total
        total_codigo_servico = {
            'Código do Serviço': 'TOTAL',
            'Quantidade de Notas': codigo_servico_summary['Quantidade de Notas'].sum(),
            'Valor do Serviço (R$)': codigo_servico_summary['Valor do Serviço (R$)'].sum().round(2),
            'Base de Cálculo (R$)': codigo_servico_summary['Base de Cálculo (R$)'].sum().round(2),
            'ISS Próprio (R$)': codigo_servico_summary['ISS Próprio (R$)'].sum().round(2),
            'ISS Retido (R$)': codigo_servico_summary['ISS Retido (R$)'].sum().round(2),
            'Alíquota (%)': None  # Não há alíquota para o total
        }

        codigo_servico_summary = pd.concat([codigo_servico_summary, pd.DataFrame([total_codigo_servico])])

        # Adicionar cabeçalhos
        headers = ['Código do Serviço', 'Quantidade de Notas', 'Valor do Serviço (R$)',
                'Base de Cálculo (R$)', 'Alíquota (%)', 'ISS Próprio (R$)', 'ISS Retido (R$)']

        # Definir larguras das colunas para a tabela de código do serviço
        column_widths = {
            'A': 20,  # Código do Serviço
            'B': 15,  # Quantidade de Notas
            'C': 20,  # Valor do Serviço
            'D': 20,  # Base de Cálculo
            'E': 15,  # Alíquota
            'F': 15,  # ISS Próprio
            'G': 15,  # ISS Retido
        }

        for col_letter, width in column_widths.items():
            worksheet.column_dimensions[col_letter].width = width

        for i, header in enumerate(headers):
            cell = worksheet.cell(row=current_row, column=i+1)
            cell.value = header
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
            cell.alignment = Alignment(horizontal='center')
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )

        current_row += 1

        # Adicionar dados
        for _, row in codigo_servico_summary.iterrows():
            for i, col in enumerate(headers):
                cell = worksheet.cell(row=current_row, column=i+1)

                # Formatação especial para valores monetários
                if 'R$' in col:
                    # Armazenar o valor como número, não como string, e arredondar para 2 casas decimais
                    try:
                        cell.value = round(float(row[col]), 2)
                    except (ValueError, TypeError):
                        # Se não for possível converter para float, usar o valor original
                        if isinstance(row[col], (int, float)):
                            cell.value = row[col]
                        else:
                            # Tentar converter para float novamente
                            try:
                                cell.value = float(str(row[col]).replace('R$', '').replace('.', '').replace(',', '.').strip())
                            except:
                                cell.value = 0.0
                    # Usar formato de moeda simples (código 44 no Excel)
                    cell.number_format = 'R$ #,##0.00'
                    cell.alignment = Alignment(horizontal='right')
                elif col == 'Alíquota (%)':
                    # Verificar se o valor não é None
                    if pd.notna(row[col]):
                        # Obter o valor e garantir que seja um número
                        value = float(row[col])
                        # Se o valor for muito alto (ex: 300), dividir por 100 para obter 3
                        if value > 100:
                            value = value / 100
                        cell.value = value
                        # Usar formato numérico com 2 casas decimais seguido de %
                        cell.number_format = '0.00"%"'
                    cell.alignment = Alignment(horizontal='center')
                elif col == 'Quantidade de Notas':
                    cell.value = int(row[col])
                    cell.alignment = Alignment(horizontal='center')
                elif col == 'Código do Serviço':
                    # Formatar o código do serviço (ex: 0403 para 4.3)
                    if row[col] != 'TOTAL':  # Não formatar a linha de total
                        cell.value = format_codigo_servico(row[col])
                    else:
                        cell.value = row[col]
                    cell.alignment = Alignment(horizontal='center')
                else:
                    cell.value = row[col]

                # Formatação especial para a linha de total
                if row['Código do Serviço'] == 'TOTAL':
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")

                # Adicionar bordas
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )

            current_row += 1

        # Espaço entre tabelas
        current_row += 2

    # ===== TABELA 4: TOP 5 MAIORES VALORES POR ANO =====
    # Adicionar título da seção
    worksheet.merge_cells(f'A{current_row}:E{current_row}')
    section_title = worksheet[f'A{current_row}']
    section_title.value = "TOP 5 NOTAS FISCAIS DE MAIOR VALOR POR ANO (ESCRITURADAS E QUITADAS)"
    section_title.font = Font(bold=True, size=12)
    section_title.alignment = Alignment(horizontal='center', vertical='center')
    section_title.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    section_title.font = Font(bold=True, size=12, color="FFFFFF")

    # Definir altura da linha do título da seção
    worksheet.row_dimensions[current_row].height = 20

    current_row += 1

    # Filtrar apenas notas escrituradas e quitadas
    df_validadas = df[df['Situação'].isin(['ESCRITURADA', 'QUITADA'])]

    # Verificar se temos a coluna de competência
    if 'Competência' in df_validadas.columns:
        try:
            # Extrair o ano da competência (formato MM/AAAA)
            df_validadas['temp_date'] = pd.to_datetime(
                df_validadas['Competência'],
                format='%m/%Y',
                errors='coerce'
            )
            df_validadas['Ano'] = df_validadas['temp_date'].dt.year

            # Obter os anos únicos e ordená-los
            anos = sorted(df_validadas['Ano'].dropna().unique())

            # Obter as top 5 notas de cada ano
            top_notas_por_ano = []

            for ano in anos:
                # Filtrar notas do ano atual
                df_ano = df_validadas[df_validadas['Ano'] == ano]

                # Obter as 5 notas de maior valor
                top_notas_ano = df_ano.sort_values('Valor do Serviço (R$)', ascending=False).head(5)

                if not top_notas_ano.empty:
                    # Adicionar uma coluna para identificar o ano (para formatação)
                    top_notas_ano['Grupo_Ano'] = ano
                    top_notas_por_ano.append(top_notas_ano)

            # Combinar todas as notas de maior valor por ano
            if top_notas_por_ano:
                top_notas = pd.concat(top_notas_por_ano)

                # Ordenar por ano (decrescente) e valor do serviço (decrescente)
                top_notas = top_notas.sort_values(['Ano', 'Valor do Serviço (R$)'], ascending=[False, False])

                # Remover as colunas temporárias que não serão exibidas
                if 'temp_date' in top_notas.columns:
                    top_notas = top_notas.drop('temp_date', axis=1)

                # Garantir que os valores monetários estão arredondados para 2 casas decimais
                if 'Valor do Serviço (R$)' in top_notas.columns:
                    top_notas['Valor do Serviço (R$)'] = top_notas['Valor do Serviço (R$)'].round(2)
            else:
                # Se não houver notas, usar o método original
                top_notas = df_validadas.sort_values('Valor do Serviço (R$)', ascending=False).head(5)

                # Garantir que os valores monetários estão arredondados para 2 casas decimais
                if 'Valor do Serviço (R$)' in top_notas.columns:
                    top_notas['Valor do Serviço (R$)'] = top_notas['Valor do Serviço (R$)'].round(2)
        except Exception as e:
            logger.warning(f"Erro ao processar top 5 por ano: {e}")
            # Se falhar, usar o método original
            top_notas = df_validadas.sort_values('Valor do Serviço (R$)', ascending=False).head(5)

            # Garantir que os valores monetários estão arredondados para 2 casas decimais
            if 'Valor do Serviço (R$)' in top_notas.columns:
                top_notas['Valor do Serviço (R$)'] = top_notas['Valor do Serviço (R$)'].round(2)
    else:
        # Se não houver coluna de competência, usar o método original
        top_notas = df_validadas.sort_values('Valor do Serviço (R$)', ascending=False).head(5)

        # Garantir que os valores monetários estão arredondados para 2 casas decimais
        if 'Valor do Serviço (R$)' in top_notas.columns:
            top_notas['Valor do Serviço (R$)'] = top_notas['Valor do Serviço (R$)'].round(2)

    # Selecionar colunas relevantes
    top_columns = [
        'Número da Nota', 'Competência', 'Nome do Tomador',
        'Valor do Serviço (R$)', 'Situação'
    ]

    # Verificar se todas as colunas existem
    missing_top_columns = [col for col in top_columns if col not in df.columns]
    if missing_top_columns:
        # Usar apenas as colunas disponíveis
        top_columns = [col for col in top_columns if col in df.columns]
        top_columns.append('Valor do Serviço (R$)')  # Garantir que esta coluna esteja presente

    # Definir larguras das colunas para a tabela de top 5
    column_widths = {
        'A': 20,  # Número da Nota
        'B': 15,  # Competência
        'C': 40,  # Nome do Tomador
        'D': 20,  # Valor do Serviço
        'E': 15,  # Situação
    }

    for col_letter, width in column_widths.items():
        worksheet.column_dimensions[col_letter].width = width

    # Adicionar cabeçalhos
    for i, header in enumerate(top_columns):
        cell = worksheet.cell(row=current_row, column=i+1)
        cell.value = header
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        cell.alignment = Alignment(horizontal='center')
        cell.border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )

    current_row += 1

    # Adicionar dados
    # Verificar se temos a coluna de ano para agrupar
    if 'Grupo_Ano' in top_notas.columns:
        # Obter os anos únicos e ordená-los (decrescente)
        anos_unicos = sorted(top_notas['Grupo_Ano'].unique(), reverse=True)

        # Para cada ano, adicionar um cabeçalho e as notas correspondentes
        for ano in anos_unicos:
            # Adicionar cabeçalho do ano
            worksheet.merge_cells(f'A{current_row}:E{current_row}')
            year_header = worksheet[f'A{current_row}']
            year_header.value = f"TOP 5 NOTAS FISCAIS DE {int(ano)}"
            year_header.font = Font(bold=True)
            year_header.alignment = Alignment(horizontal='center', vertical='center')
            year_header.fill = PatternFill(start_color="E6F0FF", end_color="E6F0FF", fill_type="solid")

            # Adicionar bordas ao cabeçalho
            for col_idx in range(1, 6):  # 5 colunas
                cell = worksheet.cell(row=current_row, column=col_idx)
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )

            current_row += 1

            # Filtrar notas do ano atual
            notas_ano = top_notas[top_notas['Grupo_Ano'] == ano]

            # Adicionar as notas do ano
            for i, (_, row) in enumerate(notas_ano.iterrows()):
                # Definir cor de fundo alternada
                row_color = "F5F5F5" if i % 2 == 0 else "FFFFFF"  # Cinza claro / Branco

                for j, col in enumerate(top_columns):
                    cell = worksheet.cell(row=current_row, column=j+1)

                    # Formatação especial para valores monetários
                    if 'R$' in col:
                        # Armazenar o valor como número, não como string, e arredondar para 2 casas decimais
                        try:
                            cell.value = round(float(row[col]), 2)
                        except (ValueError, TypeError):
                            # Se não for possível converter para float, usar o valor original
                            if isinstance(row[col], (int, float)):
                                cell.value = row[col]
                            else:
                                # Tentar converter para float novamente
                                try:
                                    cell.value = float(str(row[col]).replace('R$', '').replace('.', '').replace(',', '.').strip())
                                except:
                                    cell.value = 0.0
                        # Usar formato de moeda simples (código 44 no Excel)
                        cell.number_format = 'R$ #,##0.00'
                        cell.alignment = Alignment(horizontal='right', vertical='center')
                    # Formatação especial para competência
                    elif col == 'Competência' and pd.notna(row[col]):
                        cell.value = row[col]
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    elif col == 'Número da Nota':
                        cell.value = row[col]
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    elif col == 'Situação':
                        cell.value = row[col]
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    elif col == 'Nome do Tomador':
                        # Limitar o tamanho do nome do tomador para evitar nomes muito longos
                        nome_tomador = row[col]
                        if len(nome_tomador) > 35:
                            nome_tomador = nome_tomador[:32] + "..."
                        cell.value = nome_tomador
                        cell.alignment = Alignment(horizontal='left', vertical='center')
                    elif col == 'Alíquota (%)':
                        if pd.notna(row[col]):
                            value = float(row[col])
                            # Se o valor for muito alto (ex: 300), dividir por 100 para obter 3
                            if value > 100:
                                value = value / 100
                            cell.value = value
                            # Usar formato numérico com 2 casas decimais seguido de %
                            cell.number_format = '0.00"%"'
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    else:
                        cell.value = row[col]
                        cell.alignment = Alignment(vertical='center')

                    # Aplicar cor de fundo
                    cell.fill = PatternFill(start_color=row_color, end_color=row_color, fill_type="solid")

                    # Adicionar bordas
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )

                current_row += 1

            # Adicionar espaço entre anos
            current_row += 1
    else:
        # Se não temos a coluna de ano, usar o formato original
        for i, (_, row) in enumerate(top_notas[top_columns].iterrows()):
            # Definir cor de fundo alternada
            row_color = "F5F5F5" if i % 2 == 0 else "FFFFFF"  # Cinza claro / Branco

            for j, col in enumerate(top_columns):
                cell = worksheet.cell(row=current_row, column=j+1)

                # Formatação especial para valores monetários
                if 'R$' in col:
                    # Armazenar o valor como número, não como string, e arredondar para 2 casas decimais
                    try:
                        cell.value = round(float(row[col]), 2)
                    except (ValueError, TypeError):
                        # Se não for possível converter para float, usar o valor original
                        if isinstance(row[col], (int, float)):
                            cell.value = row[col]
                        else:
                            # Tentar converter para float novamente
                            try:
                                cell.value = float(str(row[col]).replace('R$', '').replace('.', '').replace(',', '.').strip())
                            except:
                                cell.value = 0.0
                    # Usar formato de moeda simples (código 44 no Excel)
                    cell.number_format = 'R$ #,##0.00'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                # Formatação especial para competência
                elif col == 'Competência' and pd.notna(row[col]):
                    cell.value = row[col]
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                elif col == 'Número da Nota':
                    cell.value = row[col]
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                elif col == 'Situação':
                    cell.value = row[col]
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                elif col == 'Nome do Tomador':
                    # Limitar o tamanho do nome do tomador para evitar nomes muito longos
                    nome_tomador = row[col]
                    if len(nome_tomador) > 35:
                        nome_tomador = nome_tomador[:32] + "..."
                    cell.value = nome_tomador
                    cell.alignment = Alignment(horizontal='left', vertical='center')
                elif col == 'Alíquota (%)':
                    if pd.notna(row[col]):
                        value = float(row[col])
                        # Se o valor for muito alto (ex: 300), dividir por 100 para obter 3
                        if value > 100:
                            value = value / 100
                        cell.value = value
                        # Usar formato numérico com 2 casas decimais seguido de %
                        cell.number_format = '0.00"%"'
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    cell.value = row[col]
                    cell.alignment = Alignment(vertical='center')

                # Aplicar cor de fundo
                cell.fill = PatternFill(start_color=row_color, end_color=row_color, fill_type="solid")

                # Adicionar bordas
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )

            current_row += 1

    # ===== TABELA 5: TOP 10 TOMADORES =====
    # Espaço entre tabelas
    current_row += 3

    # Adicionar título da seção
    worksheet.merge_cells(f'A{current_row}:D{current_row}')
    section_title = worksheet[f'A{current_row}']
    section_title.value = "TOP 10 TOMADORES POR VALOR TOTAL"
    section_title.font = Font(bold=True, size=12)
    section_title.alignment = Alignment(horizontal='center', vertical='center')
    section_title.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    section_title.font = Font(bold=True, size=12, color="FFFFFF")

    # Definir altura da linha do título da seção
    worksheet.row_dimensions[current_row].height = 20

    current_row += 1

    # Filtrar apenas notas escrituradas e quitadas
    df_validadas = df[df['Situação'].isin(['ESCRITURADA', 'QUITADA'])]

    # Verificar se temos a coluna de tomador
    if 'Nome do Tomador' in df_validadas.columns:
        # Agrupar por tomador e calcular o valor total e a quantidade de notas
        tomadores_summary = df_validadas.groupby('Nome do Tomador').agg({
            'Número da Nota': 'count',
            'Valor do Serviço (R$)': 'sum'
        }).reset_index()

        # Renomear as colunas
        tomadores_summary = tomadores_summary.rename(columns={
            'Número da Nota': 'Quantidade de Notas',
            'Valor do Serviço (R$)': 'Valor Total (R$)'
        })

        # Ordenar por valor total (decrescente) e selecionar os 10 primeiros
        tomadores_summary = tomadores_summary.sort_values('Valor Total (R$)', ascending=False).head(10)

        # Arredondar valores para 2 casas decimais
        tomadores_summary['Valor Total (R$)'] = tomadores_summary['Valor Total (R$)'].round(2)

        # Calcular o valor médio por nota
        tomadores_summary['Valor Médio por Nota (R$)'] = (
            tomadores_summary['Valor Total (R$)'] / tomadores_summary['Quantidade de Notas']
        ).round(2)

        # Definir larguras das colunas para a tabela de top 10 tomadores
        column_widths = {
            'A': 45,  # Nome do Tomador
            'B': 18,  # Quantidade de Notas
            'C': 22,  # Valor Total
            'D': 22,  # Valor Médio por Nota
        }

        for col_letter, width in column_widths.items():
            worksheet.column_dimensions[col_letter].width = width

        # Adicionar cabeçalhos
        headers = [
            'Nome do Tomador', 'Quantidade de Notas', 'Valor Total (R$)',
            'Valor Médio por Nota (R$)'
        ]

        for i, header in enumerate(headers):
            cell = worksheet.cell(row=current_row, column=i+1)
            cell.value = header
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
            cell.alignment = Alignment(horizontal='center')
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )

        current_row += 1

        # Adicionar dados
        total_notas = 0
        total_valor = 0

        for i, (_, row) in enumerate(tomadores_summary.iterrows()):
            # Destacar o primeiro tomador (maior valor)
            highlight = (i == 0)

            # Alternar cores de fundo para melhorar a legibilidade
            row_color = "F5F5F5" if i % 2 == 0 else "FFFFFF"

            for j, col in enumerate(['Nome do Tomador', 'Quantidade de Notas', 'Valor Total (R$)', 'Valor Médio por Nota (R$)']):
                cell = worksheet.cell(row=current_row, column=j+1)

                if col == 'Nome do Tomador':
                    # Limitar o tamanho do nome do tomador para evitar nomes muito longos
                    nome_tomador = row[col]
                    if len(nome_tomador) > 40:
                        nome_tomador = nome_tomador[:37] + "..."
                    cell.value = nome_tomador
                    cell.alignment = Alignment(horizontal='left', vertical='center')
                elif col == 'Quantidade de Notas':
                    cell.value = int(row[col])
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    total_notas += int(row[col])
                elif col == 'Valor Total (R$)':
                    # Armazenar o valor como número, não como string
                    cell.value = float(row[col])
                    # Usar formato de moeda
                    cell.number_format = 'R$ #,##0.00'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                    total_valor += float(row[col])
                elif col == 'Valor Médio por Nota (R$)':
                    # Armazenar o valor como número, não como string
                    cell.value = float(row[col])
                    # Usar formato de moeda
                    cell.number_format = 'R$ #,##0.00'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                else:
                    cell.value = row[col]

                # Destacar o primeiro tomador (maior valor)
                if highlight:
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="E6F0FF", end_color="E6F0FF", fill_type="solid")
                else:
                    # Aplicar cor alternada para as outras linhas
                    cell.fill = PatternFill(start_color=row_color, end_color=row_color, fill_type="solid")

                # Adicionar bordas
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )

            current_row += 1

        # Adicionar linha de subtotal
        for j, col in enumerate(['Total Top 10', total_notas, total_valor, '']):
            cell = worksheet.cell(row=current_row, column=j+1)

            if j == 0:
                cell.value = col
                cell.alignment = Alignment(horizontal='left', vertical='center')
            elif j == 1:
                cell.value = col
                cell.alignment = Alignment(horizontal='center', vertical='center')
            elif j == 2:
                cell.value = col
                cell.number_format = 'R$ #,##0.00'
                cell.alignment = Alignment(horizontal='right', vertical='center')

            # Formatação da linha de total
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")

            # Adicionar bordas
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )

        current_row += 1

    # ===== TABELA 6: TOP 5 MESES COM MAIOR FATURAMENTO =====
    # Espaço entre tabelas
    current_row += 3

    # Adicionar título da seção
    worksheet.merge_cells(f'A{current_row}:E{current_row}')
    section_title = worksheet[f'A{current_row}']
    section_title.value = "TOP 5 MESES COM MAIOR FATURAMENTO"
    section_title.font = Font(bold=True, size=12)
    section_title.alignment = Alignment(horizontal='center', vertical='center')
    section_title.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    section_title.font = Font(bold=True, size=12, color="FFFFFF")

    # Definir altura da linha do título da seção
    worksheet.row_dimensions[current_row].height = 20

    current_row += 1

    # Filtrar apenas notas escrituradas e quitadas
    df_validadas = df[df['Situação'].isin(['ESCRITURADA', 'QUITADA'])]

    # Verificar se temos as colunas necessárias
    if 'Mês' in df_validadas.columns:
        # Criar coluna de mês por extenso
        meses_nomes = {
            1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril', 5: 'Maio', 6: 'Junho',
            7: 'Julho', 8: 'Agosto', 9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
        }
        df_validadas['Mês Nome'] = df_validadas['Mês'].map(meses_nomes)

        # Agrupar por mês e calcular o valor total
        meses_summary = df_validadas.groupby(['Mês', 'Mês Nome']).agg({
            'Número da Nota': 'count',
            'Valor do Serviço (R$)': 'sum'
        }).reset_index()

        # Renomear as colunas
        meses_summary = meses_summary.rename(columns={
            'Número da Nota': 'Quantidade de Notas',
            'Valor do Serviço (R$)': 'Valor Total (R$)'
        })

        # Ordenar por valor total (decrescente) e selecionar os 5 primeiros
        meses_summary = meses_summary.sort_values('Valor Total (R$)', ascending=False).head(5)

        # Calcular o valor médio por nota
        meses_summary['Valor Médio por Nota (R$)'] = (
            meses_summary['Valor Total (R$)'] / meses_summary['Quantidade de Notas']
        ).round(2)

        # Definir larguras das colunas
        column_widths = {
            'A': 15,  # Mês (número)
            'B': 15,  # Mês Nome
            'C': 18,  # Quantidade de Notas
            'D': 22,  # Valor Total
            'E': 22,  # Valor Médio por Nota
        }

        for col_letter, width in column_widths.items():
            worksheet.column_dimensions[col_letter].width = width

        # Adicionar cabeçalhos
        headers = [
            'Mês', 'Mês Nome', 'Quantidade de Notas', 'Valor Total (R$)',
            'Valor Médio por Nota (R$)'
        ]

        for i, header in enumerate(headers):
            cell = worksheet.cell(row=current_row, column=i+1)
            cell.value = header
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
            cell.alignment = Alignment(horizontal='center')
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )

        current_row += 1

        # Adicionar dados
        for i, (_, row) in enumerate(meses_summary.iterrows()):
            # Alternar cores de fundo para melhorar a legibilidade
            row_color = "F5F5F5" if i % 2 == 0 else "FFFFFF"

            for j, col in enumerate(['Mês', 'Mês Nome', 'Quantidade de Notas', 'Valor Total (R$)', 'Valor Médio por Nota (R$)']):
                cell = worksheet.cell(row=current_row, column=j+1)
                
                if col == 'Mês':
                    cell.value = int(row[col])
                    cell.alignment = Alignment(horizontal='center')
                elif col == 'Mês Nome':
                    cell.value = row[col]
                    cell.alignment = Alignment(horizontal='left')
                elif col == 'Quantidade de Notas':
                    cell.value = int(row[col])
                    cell.alignment = Alignment(horizontal='center')
                elif col in ['Valor Total (R$)', 'Valor Médio por Nota (R$)']:
                    cell.value = float(row[col])
                    cell.number_format = 'R$ #,##0.00'
                    cell.alignment = Alignment(horizontal='right')

                # Aplicar cor de fundo
                cell.fill = PatternFill(start_color=row_color, end_color=row_color, fill_type="solid")

                # Adicionar bordas
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )

            current_row += 1

    # ===== TABELA 7: TOP 5 CÓDIGOS DE SERVIÇO MAIS RENTÁVEIS =====
    # Espaço entre tabelas
    current_row += 3

    # Adicionar título da seção
    worksheet.merge_cells(f'A{current_row}:E{current_row}')
    section_title = worksheet[f'A{current_row}']
    section_title.value = "TOP 5 CÓDIGOS DE SERVIÇO MAIS RENTÁVEIS"
    section_title.font = Font(bold=True, size=12)
    section_title.alignment = Alignment(horizontal='center', vertical='center')
    section_title.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    section_title.font = Font(bold=True, size=12, color="FFFFFF")

    # Definir altura da linha do título da seção
    worksheet.row_dimensions[current_row].height = 20

    current_row += 1

    # Verificar se temos a coluna de código do serviço
    if 'Código do Serviço' in df_validadas.columns:
        # Agrupar por código do serviço e calcular o valor total e a quantidade de notas
        servicos_summary = df_validadas.groupby('Código do Serviço').agg({
            'Número da Nota': 'count',
            'Valor do Serviço (R$)': 'sum'
        }).reset_index()

        # Renomear as colunas
        servicos_summary = servicos_summary.rename(columns={
            'Número da Nota': 'Quantidade de Notas',
            'Valor do Serviço (R$)': 'Valor Total (R$)'
        })

        # Calcular o valor médio por nota
        servicos_summary['Valor Médio por Nota (R$)'] = (
            servicos_summary['Valor Total (R$)'] / servicos_summary['Quantidade de Notas']
        ).round(2)

        # Ordenar por valor médio por nota (decrescente) e selecionar os 5 primeiros
        servicos_summary = servicos_summary.sort_values('Valor Médio por Nota (R$)', ascending=False).head(5)

        # Definir larguras das colunas
        column_widths = {
            'A': 20,  # Código do Serviço
            'B': 18,  # Quantidade de Notas
            'C': 22,  # Valor Total
            'D': 22,  # Valor Médio por Nota
        }

        for col_letter, width in column_widths.items():
            worksheet.column_dimensions[col_letter].width = width

        # Adicionar cabeçalhos
        headers = [
            'Código do Serviço', 'Quantidade de Notas', 'Valor Total (R$)',
            'Valor Médio por Nota (R$)'
        ]

        for i, header in enumerate(headers):
            cell = worksheet.cell(row=current_row, column=i+1)
            cell.value = header
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
            cell.alignment = Alignment(horizontal='center')
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )

        current_row += 1

        # Adicionar dados
        for i, (_, row) in enumerate(servicos_summary.iterrows()):
            # Alternar cores de fundo para melhorar a legibilidade
            row_color = "F5F5F5" if i % 2 == 0 else "FFFFFF"

            for j, col in enumerate(['Código do Serviço', 'Quantidade de Notas', 'Valor Total (R$)', 'Valor Médio por Nota (R$)']):
                cell = worksheet.cell(row=current_row, column=j+1)
                
                if col == 'Código do Serviço':
                    cell.value = str(row[col])
                    cell.alignment = Alignment(horizontal='left')
                elif col == 'Quantidade de Notas':
                    cell.value = int(row[col])
                    cell.alignment = Alignment(horizontal='center')
                elif col in ['Valor Total (R$)', 'Valor Médio por Nota (R$)']:
                    cell.value = float(row[col])
                    cell.number_format = 'R$ #,##0.00'
                    cell.alignment = Alignment(horizontal='right')

                # Aplicar cor de fundo
                cell.fill = PatternFill(start_color=row_color, end_color=row_color, fill_type="solid")

                # Adicionar bordas
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )

            current_row += 1

    # ===== TABELA 8: TOP 5 TOMADORES COM MAIOR CRESCIMENTO =====
    # Espaço entre tabelas
    current_row += 3

    # Adicionar título da seção
    worksheet.merge_cells(f'A{current_row}:E{current_row}')
    section_title = worksheet[f'A{current_row}']
    section_title.value = "TOP 5 TOMADORES COM MAIOR CRESCIMENTO ENTRE PERÍODOS"
    section_title.font = Font(bold=True, size=12)
    section_title.alignment = Alignment(horizontal='center', vertical='center')
    section_title.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    section_title.font = Font(bold=True, size=12, color="FFFFFF")

    # Definir altura da linha do título da seção
    worksheet.row_dimensions[current_row].height = 20

    current_row += 1

    # Verificar se temos as colunas necessárias
    if 'Nome do Tomador' in df_validadas.columns and 'Mês' in df_validadas.columns and 'Ano' in df_validadas.columns:
        # Criar coluna de semestre
        df_validadas['Semestre'] = ((df_validadas['Mês'] - 1) // 6) + 1
        df_validadas['Período'] = df_validadas['Ano'].astype(str) + '-S' + df_validadas['Semestre'].astype(str)
        
        # Obter os dois períodos mais recentes
        periodos = sorted(df_validadas['Período'].unique())
        if len(periodos) >= 2:
            periodo_atual = periodos[-1]
            periodo_anterior = periodos[-2]
            
            # Agrupar por tomador e período e calcular o valor total
            tomadores_periodos = df_validadas[df_validadas['Período'].isin([periodo_anterior, periodo_atual])].groupby(
                ['Nome do Tomador', 'Período']
            ).agg({
                'Valor do Serviço (R$)': 'sum'
            }).reset_index()
            
            # Criar um pivot para ter uma coluna para cada período
            tomadores_pivot = tomadores_periodos.pivot(
                index='Nome do Tomador',
                columns='Período',
                values='Valor do Serviço (R$)'
            ).reset_index()
            
            # Renomear as colunas
            tomadores_pivot = tomadores_pivot.rename(columns={
                periodo_anterior: 'Valor Período Anterior (R$)',
                periodo_atual: 'Valor Período Atual (R$)'
            })
            
            # Filtrar apenas tomadores que aparecem em ambos os períodos e têm valor anterior > 0
            tomadores_pivot = tomadores_pivot.dropna()
            tomadores_pivot = tomadores_pivot[tomadores_pivot['Valor Período Anterior (R$)'] > 0]
            
            # Calcular o crescimento percentual
            tomadores_pivot['Crescimento (%)'] = (
                (tomadores_pivot['Valor Período Atual (R$)'] - tomadores_pivot['Valor Período Anterior (R$)']) / 
                tomadores_pivot['Valor Período Anterior (R$)'] * 100
            ).round(2)
            
            # Ordenar por crescimento (decrescente) e selecionar os 5 primeiros
            tomadores_pivot = tomadores_pivot.sort_values('Crescimento (%)', ascending=False).head(5)
            
            # Definir larguras das colunas
            column_widths = {
                'A': 40,  # Nome do Tomador
                'B': 25,  # Valor Período Anterior
                'C': 25,  # Valor Período Atual
                'D': 20,  # Crescimento
            }

            for col_letter, width in column_widths.items():
                worksheet.column_dimensions[col_letter].width = width

            # Adicionar cabeçalhos
            headers = [
                'Nome do Tomador', 'Valor Período Anterior (R$)',
                'Valor Período Atual (R$)', 'Crescimento (%)'
            ]

            # Adicionar informação sobre os períodos
            periodo_info = worksheet.cell(row=current_row, column=1)
            periodo_info.value = f"Períodos comparados: {periodo_anterior} vs {periodo_atual}"
            periodo_info.font = Font(italic=True)
            worksheet.merge_cells(f'A{current_row}:D{current_row}')
            
            current_row += 1

            for i, header in enumerate(headers):
                cell = worksheet.cell(row=current_row, column=i+1)
                cell.value = header
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
                cell.alignment = Alignment(horizontal='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )

            current_row += 1

            # Adicionar dados
            for i, (_, row) in enumerate(tomadores_pivot.iterrows()):
                # Alternar cores de fundo para melhorar a legibilidade
                row_color = "F5F5F5" if i % 2 == 0 else "FFFFFF"

                for j, col in enumerate(['Nome do Tomador', 'Valor Período Anterior (R$)', 'Valor Período Atual (R$)', 'Crescimento (%)']):
                    cell = worksheet.cell(row=current_row, column=j+1)
                    
                    if col == 'Nome do Tomador':
                        nome_tomador = row[col]
                        if len(nome_tomador) > 35:
                            nome_tomador = nome_tomador[:32] + "..."
                        cell.value = nome_tomador
                        cell.alignment = Alignment(horizontal='left')
                    elif col in ['Valor Período Anterior (R$)', 'Valor Período Atual (R$)']:
                        cell.value = float(row[col])
                        cell.number_format = 'R$ #,##0.00'
                        cell.alignment = Alignment(horizontal='right')
                    elif col == 'Crescimento (%)':
                        cell.value = float(row[col])
                        cell.number_format = '0.00"%"'
                        cell.alignment = Alignment(horizontal='center')
                        # Destacar crescimento positivo em verde e negativo em vermelho
                        if row[col] > 0:
                            cell.font = Font(color="006100")
                        elif row[col] < 0:
                            cell.font = Font(color="9C0006")

                    # Aplicar cor de fundo
                    cell.fill = PatternFill(start_color=row_color, end_color=row_color, fill_type="solid")

                    # Adicionar bordas
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )

                current_row += 1

    # ===== TABELA 9: TOP 5 DIAS DO MÊS COM MAIOR EMISSÃO DE NOTAS =====
    # Espaço entre tabelas
    current_row += 3

    # Adicionar título da seção
    worksheet.merge_cells(f'A{current_row}:D{current_row}')
    section_title = worksheet[f'A{current_row}']
    section_title.value = "TOP 5 DIAS DO MÊS COM MAIOR EMISSÃO DE NOTAS"
    section_title.font = Font(bold=True, size=12)
    section_title.alignment = Alignment(horizontal='center', vertical='center')
    section_title.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    section_title.font = Font(bold=True, size=12, color="FFFFFF")

    # Definir altura da linha do título da seção
    worksheet.row_dimensions[current_row].height = 20

    current_row += 1

    # Verificar se temos a coluna de data de emissão
    if 'Data de Emissão' in df.columns:
        # Extrair o dia do mês da data de emissão
        df['Dia do Mês'] = pd.to_datetime(df['Data de Emissão'], errors='coerce').dt.day
        
        # Agrupar por dia do mês e calcular a quantidade de notas e o valor total
        dias_summary = df.groupby('Dia do Mês').agg({
            'Número da Nota': 'count',
            'Valor do Serviço (R$)': 'sum'
        }).reset_index()
        
        # Renomear as colunas
        dias_summary = dias_summary.rename(columns={
            'Número da Nota': 'Quantidade de Notas',
            'Valor do Serviço (R$)': 'Valor Total (R$)'
        })
        
        # Ordenar por quantidade de notas (decrescente) e selecionar os 5 primeiros
        dias_summary = dias_summary.sort_values('Quantidade de Notas', ascending=False).head(5)
        
        # Calcular o valor médio por nota
        dias_summary['Valor Médio por Nota (R$)'] = (
            dias_summary['Valor Total (R$)'] / dias_summary['Quantidade de Notas']
        ).round(2)
        
        # Definir larguras das colunas
        column_widths = {
            'A': 15,  # Dia do Mês
            'B': 18,  # Quantidade de Notas
            'C': 22,  # Valor Total
            'D': 22,  # Valor Médio por Nota
        }

        for col_letter, width in column_widths.items():
            worksheet.column_dimensions[col_letter].width = width

        # Adicionar cabeçalhos
        headers = [
            'Dia do Mês', 'Quantidade de Notas', 'Valor Total (R$)',
            'Valor Médio por Nota (R$)'
        ]

        for i, header in enumerate(headers):
            cell = worksheet.cell(row=current_row, column=i+1)
            cell.value = header
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
            cell.alignment = Alignment(horizontal='center')
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )

        current_row += 1

        # Adicionar dados
        for i, (_, row) in enumerate(dias_summary.iterrows()):
            # Alternar cores de fundo para melhorar a legibilidade
            row_color = "F5F5F5" if i % 2 == 0 else "FFFFFF"

            for j, col in enumerate(['Dia do Mês', 'Quantidade de Notas', 'Valor Total (R$)', 'Valor Médio por Nota (R$)']):
                cell = worksheet.cell(row=current_row, column=j+1)
                
                if col == 'Dia do Mês':
                    cell.value = int(row[col])
                    cell.alignment = Alignment(horizontal='center')
                elif col == 'Quantidade de Notas':
                    cell.value = int(row[col])
                    cell.alignment = Alignment(horizontal='center')
                elif col in ['Valor Total (R$)', 'Valor Médio por Nota (R$)']:
                    cell.value = float(row[col])
                    cell.number_format = 'R$ #,##0.00'
                    cell.alignment = Alignment(horizontal='right')

                # Aplicar cor de fundo
                cell.fill = PatternFill(start_color=row_color, end_color=row_color, fill_type="solid")

                # Adicionar bordas
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )

            current_row += 1

    # ===== TABELA 10: TOP 5 COMPETÊNCIAS COM MAIOR VALOR DE ISS =====
    # Espaço entre tabelas
    current_row += 3

    # Adicionar título da seção
    worksheet.merge_cells(f'A{current_row}:E{current_row}')
    section_title = worksheet[f'A{current_row}']
    section_title.value = "TOP 5 COMPETÊNCIAS COM MAIOR VALOR DE ISS"
    section_title.font = Font(bold=True, size=12)
    section_title.alignment = Alignment(horizontal='center', vertical='center')
    section_title.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    section_title.font = Font(bold=True, size=12, color="FFFFFF")

    # Definir altura da linha do título da seção
    worksheet.row_dimensions[current_row].height = 20

    current_row += 1

    # Verificar se temos as colunas necessárias
    if 'Competência' in df_validadas.columns and 'ISS Próprio (R$)' in df_validadas.columns and 'ISS Retido (R$)' in df_validadas.columns:
        # Calcular o ISS total para cada nota
        df_validadas['ISS Total (R$)'] = df_validadas['ISS Próprio (R$)'] + df_validadas['ISS Retido (R$)']
        
        # Agrupar por competência e calcular o valor total de ISS
        competencias_summary = df_validadas.groupby('Competência').agg({
            'Número da Nota': 'count',
            'Valor do Serviço (R$)': 'sum',
            'ISS Total (R$)': 'sum'
        }).reset_index()
        
        # Renomear as colunas
        competencias_summary = competencias_summary.rename(columns={
            'Número da Nota': 'Quantidade de Notas',
            'Valor do Serviço (R$)': 'Valor Total (R$)'
        })
        
        # Ordenar por valor total de ISS (decrescente) e selecionar as 5 primeiras
        competencias_summary = competencias_summary.sort_values('ISS Total (R$)', ascending=False).head(5)
        
        # Calcular a alíquota média efetiva
        competencias_summary['Alíquota Média Efetiva (%)'] = (
            competencias_summary['ISS Total (R$)'] / competencias_summary['Valor Total (R$)'] * 100
        ).round(2)
        
        # Definir larguras das colunas
        column_widths = {
            'A': 15,  # Competência
            'B': 18,  # Quantidade de Notas
            'C': 22,  # Valor Total
            'D': 22,  # ISS Total
            'E': 22,  # Alíquota Média Efetiva
        }

        for col_letter, width in column_widths.items():
            worksheet.column_dimensions[col_letter].width = width

        # Adicionar cabeçalhos
        headers = [
            'Competência', 'Quantidade de Notas', 'Valor Total (R$)',
            'ISS Total (R$)', 'Alíquota Média Efetiva (%)'
        ]

        for i, header in enumerate(headers):
            cell = worksheet.cell(row=current_row, column=i+1)
            cell.value = header
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
            cell.alignment = Alignment(horizontal='center')
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )

        current_row += 1

        # Adicionar dados
        for i, (_, row) in enumerate(competencias_summary.iterrows()):
            # Alternar cores de fundo para melhorar a legibilidade
            row_color = "F5F5F5" if i % 2 == 0 else "FFFFFF"

            for j, col in enumerate(['Competência', 'Quantidade de Notas', 'Valor Total (R$)', 'ISS Total (R$)', 'Alíquota Média Efetiva (%)']):
                cell = worksheet.cell(row=current_row, column=j+1)
                
                if col == 'Competência':
                    cell.value = row[col]
                    cell.alignment = Alignment(horizontal='center')
                elif col == 'Quantidade de Notas':
                    cell.value = int(row[col])
                    cell.alignment = Alignment(horizontal='center')
                elif col in ['Valor Total (R$)', 'ISS Total (R$)']:
                    cell.value = float(row[col])
                    cell.number_format = 'R$ #,##0.00'
                    cell.alignment = Alignment(horizontal='right')
                elif col == 'Alíquota Média Efetiva (%)':
                    cell.value = float(row[col])
                    cell.number_format = '0.00"%"'
                    cell.alignment = Alignment(horizontal='center')

                # Aplicar cor de fundo
                cell.fill = PatternFill(start_color=row_color, end_color=row_color, fill_type="solid")

                # Adicionar bordas
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )

            current_row += 1

    logger.info(f"Planilha de resumo organizada criada com sucesso: {sheet_name}")



import pandas as pd
import re
from datetime import datetime
import logging

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("data_processor.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("data_processor")

def process_data(df):
    """
    Processa os dados extraídos do PDF para melhorar a estrutura e qualidade.

    Args:
        df (pd.DataFrame): DataFrame com os dados extraídos do PDF

    Returns:
        pd.DataFrame: DataFrame processado
    """
    if df.empty:
        logger.warning("DataFrame vazio. Nenhum processamento realizado.")
        return df

    logger.info("Iniciando processamento dos dados...")

    # Fazer uma cópia para não modificar o original
    processed_df = df.copy()

    # Converter tipos de dados
    processed_df = convert_data_types(processed_df)

    # Adicionar colunas calculadas
    processed_df = add_calculated_columns(processed_df)

    # Reorganizar colunas
    processed_df = reorder_columns(processed_df)

    logger.info("Processamento de dados concluído.")
    return processed_df

def clean_tomador_name(nome_tomador):
    """
    Limpa o nome do tomador, removendo códigos de serviço e valores monetários.

    Args:
        nome_tomador (str): Nome do tomador a ser limpo.

    Returns:
        str: Nome do tomador limpo.
    """
    if pd.isna(nome_tomador):
        return nome_tomador

    # Converter para string se não for
    nome_tomador = str(nome_tomador)

    # Remover código do serviço (0403) e valores monetários
    parts = nome_tomador.split()
    cleaned_parts = []

    for part in parts:
        # Pular o código do serviço
        if part == "0403" or part == "50":
            continue

        # Pular valores monetários (números com vírgula ou ponto)
        if re.match(r'^\d+[.,]\d+$', part):
            continue

        cleaned_parts.append(part)

    return ' '.join(cleaned_parts)

def convert_data_types(df):
    """
    Converte os tipos de dados das colunas para os tipos apropriados.

    Args:
        df (pd.DataFrame): DataFrame com os dados extraídos

    Returns:
        pd.DataFrame: DataFrame com tipos de dados convertidos
    """
    try:
        # Limpar nomes de tomador que podem conter código do serviço ou valores
        if 'Nome do Tomador' in df.columns:
            logger.info("Limpando nomes de tomador...")
            # Verificar se há nomes de tomador com "50 BATALHÃO"
            batalhao_count = df['Nome do Tomador'].str.contains('50 BATALHÃO', na=False).sum()
            if batalhao_count > 0:
                logger.info(f"Encontrados {batalhao_count} registros com '50 BATALHÃO' no nome do tomador")

                # Corrigir a alíquota para as linhas com BATALHÃO DE INFANTARIA
                batalhao_mask = df['Nome do Tomador'].str.contains('BATALHÃO', na=False)
                if 'Alíquota (%)' in df.columns and 'Base de Cálculo (R$)' in df.columns:
                    # Verificar se a Base de Cálculo é 3.0 e a Alíquota é 0.0
                    aliquota_mask = (df['Base de Cálculo (R$)'] == 3.0) & (df['Alíquota (%)'] == 0.0) & batalhao_mask
                    if aliquota_mask.any():
                        # Corrigir a alíquota e a base de cálculo
                        logger.info(f"Corrigindo alíquota para {aliquota_mask.sum()} linhas com BATALHÃO DE INFANTARIA")
                        df.loc[aliquota_mask, 'Alíquota (%)'] = 3.0
                        # Recalcular a base de cálculo com base no valor do serviço
                        df.loc[aliquota_mask, 'Base de Cálculo (R$)'] = df.loc[aliquota_mask, 'Valor do Serviço (R$)']

            # Limpar os nomes
            df['Nome do Tomador'] = df['Nome do Tomador'].apply(clean_tomador_name)

            # Verificar se ainda há nomes de tomador com "50 BATALHÃO" após a limpeza
            batalhao_count_after = df['Nome do Tomador'].str.contains('50 BATALHÃO', na=False).sum()
            logger.info(f"Após limpeza: {batalhao_count_after} registros com '50 BATALHÃO' no nome do tomador")

        # Converter Data de Emissão para datetime
        df['Data de Emissão'] = pd.to_datetime(df['Data de Emissão'], format='%d/%m/%Y', errors='coerce')

        # Tratar a coluna Competência (pode estar vazia em algumas linhas)
        # Manter como string para evitar problemas de conversão
        df['Competência'] = df['Competência'].astype(str)

        # Garantir que valores numéricos sejam do tipo float
        numeric_columns = [
            'Valor do Serviço (R$)',
            'Base de Cálculo (R$)',
            'Alíquota (%)',
            'ISS Próprio (R$)',
            'ISS Retido (R$)'
        ]

        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')

        # Garantir que o Número da Nota seja string
        df['Número da Nota'] = df['Número da Nota'].astype(str)

        # Formatar CPF/CNPJ como string
        df['CPF/CNPJ'] = df['CPF/CNPJ'].astype(str)

        return df

    except Exception as e:
        logger.error(f"Erro ao converter tipos de dados: {str(e)}")
        return df

def add_calculated_columns(df):
    """
    Adiciona colunas calculadas ao DataFrame.

    Args:
        df (pd.DataFrame): DataFrame com os dados extraídos

    Returns:
        pd.DataFrame: DataFrame com colunas calculadas adicionadas
    """
    try:
        # Adicionar coluna de Mês e Ano separados (para facilitar filtragem)
        # Verificar se a coluna Competência contém objetos datetime
        if pd.api.types.is_datetime64_any_dtype(df['Competência']):
            # Para linhas com datetime, extrair mês e ano
            df['Mês'] = df['Competência'].dt.month
            df['Ano'] = df['Competência'].dt.year
        else:
            # Para strings no formato "MM/YYYY", extrair mês e ano
            df['Competência'] = df['Competência'].astype(str)
            # Criar colunas vazias
            df['Mês'] = pd.NA
            df['Ano'] = pd.NA
            
            # Registrar informações sobre as competências encontradas para diagnóstico
            competencias_unicas = df['Competência'].unique()
            logger.info(f"Competências únicas encontradas: {competencias_unicas}")
            
            # Sistema robusto de reconhecimento de padrões de competência
            # Definir padrões conhecidos de competência
            padroes = [
                # Padrão padrão: MM/AAAA
                r'\b(\d{2})/(\d{4})\b',
                # Padrão com texto antes: Competência MM/AAAA
                r'[Cc]ompet[êe]ncia\s*(\d{2})/(\d{4})',
                # Padrão com texto e dois pontos: COMPETÊNCIA: MM/AAAA
                r'[Cc]ompet[êe]ncia:?\s*(\d{2})/(\d{4})',
                # Padrão com outros caracteres: MM-AAAA ou MM.AAAA
                r'\b(\d{2})[-\.](\d{4})\b'
            ]
            
            # Processar cada competência individualmente para maior precisão
            for competencia in competencias_unicas:
                if not isinstance(competencia, str) or competencia == 'nan':
                    continue
                    
                mes = None
                ano = None
                
                # Tentar cada padrão até encontrar um match
                for padrao in padroes:
                    # Verificar se o padrão tem dois grupos de captura (mês e ano)
                    if r'(\d{2})' in padrao and r'(\d{4})' in padrao:
                        match = re.search(padrao, competencia)
                        if match and len(match.groups()) >= 2:
                            mes = int(match.group(1))
                            ano = int(match.group(2))
                            logger.info(f"Competência '{competencia}' reconhecida com padrão '{padrao}': mês={mes}, ano={ano}")
                            break
                    else:
                        # Padrão com apenas um grupo de captura
                        match = re.search(padrao, competencia)
                        if match:
                            # Extrair a data completa e depois separar
                            data_completa = match.group(1)
                            partes = re.split(r'[/\-\.]', data_completa)
                            if len(partes) == 2:
                                mes = int(partes[0])
                                ano = int(partes[1])
                                logger.info(f"Competência '{competencia}' reconhecida com padrão alternativo: mês={mes}, ano={ano}")
                                break
                
                # Caso especial: tentar extrair diretamente se nenhum padrão funcionou
                if mes is None and ano is None and re.search(r'\d{2}\D\d{4}', competencia):
                    # Extrair todos os dígitos e tentar formar mês/ano
                    digitos = re.findall(r'\d+', competencia)
                    if len(digitos) >= 2:
                        # Assumir que o primeiro grupo de 1-2 dígitos é o mês e o segundo grupo de 4 dígitos é o ano
                        for i, d in enumerate(digitos):
                            if len(d) == 1 or len(d) == 2:
                                mes = int(d)
                                # Procurar o próximo grupo de 4 dígitos para o ano
                                for j in range(i+1, len(digitos)):
                                    if len(digitos[j]) == 4:
                                        ano = int(digitos[j])
                                        break
                                if ano is not None:
                                    break
                
                # Validar mês e ano
                if mes is not None and ano is not None:
                    # Verificar se o mês está no intervalo válido (1-12)
                    if 1 <= mes <= 12:
                        # Aplicar os valores encontrados ao DataFrame
                        comp_mask = df['Competência'] == competencia
                        if comp_mask.any():
                            df.loc[comp_mask, 'Mês'] = float(mes)
                            df.loc[comp_mask, 'Ano'] = float(ano)
                            logger.info(f"Atribuídos valores Mês={mes}, Ano={ano} para {comp_mask.sum()} linhas com competência '{competencia}'")
                    else:
                        logger.warning(f"Mês inválido ({mes}) encontrado na competência '{competencia}'")
            
            # Verificar se ainda existem linhas sem mês/ano atribuídos
            sem_mes_ano = df['Mês'].isna() | df['Ano'].isna()
            if sem_mes_ano.any():
                logger.warning(f"Ainda existem {sem_mes_ano.sum()} linhas sem mês/ano atribuídos após processamento")
                # Tentar extrair diretamente usando regex para essas linhas
                for idx in df[sem_mes_ano].index:
                    competencia = df.loc[idx, 'Competência']
                    if isinstance(competencia, str) and len(competencia) >= 6:  # Mínimo para ter MM/AAAA
                        # Tentar extrair qualquer padrão que pareça mês/ano
                        match = re.search(r'(\d{1,2})[^\d](\d{4})', competencia)
                        if match:
                            mes = int(match.group(1))
                            ano = int(match.group(2))
                            if 1 <= mes <= 12:
                                df.loc[idx, 'Mês'] = float(mes)
                                df.loc[idx, 'Ano'] = float(ano)
                                logger.info(f"Recuperação final: Atribuídos valores Mês={mes}, Ano={ano} para linha com competência '{competencia}'")
            
            # Registrar informações sobre mês e ano extraídos após todo o processamento
            logger.info(f"Valores únicos de mês extraídos: {df['Mês'].unique()}")
            logger.info(f"Valores únicos de ano extraídos: {df['Ano'].unique()}")
            logger.info(f"Total de linhas com mês/ano atribuídos: {(~df['Mês'].isna()).sum()} de {len(df)}")

        # Garantir que todas as colunas numéricas sejam do tipo float
        numeric_columns = [
            'Valor do Serviço (R$)', 'Base de Cálculo (R$)',
            'ISS Próprio (R$)', 'ISS Retido (R$)'
        ]

        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

        # Corrigir a alíquota para cálculos e exibição
        if 'Alíquota (%)' in df.columns:
            # Converter para numérico
            df['Alíquota (%)'] = pd.to_numeric(df['Alíquota (%)'], errors='coerce').fillna(0.0)

            # Corrigir a alíquota diretamente (se for 300, mudar para 3)
            # Isso é crucial para garantir que o valor seja armazenado como 3 em vez de 300
            mask = df['Alíquota (%)'] > 100
            df.loc[mask, 'Alíquota (%)'] = df.loc[mask, 'Alíquota (%)'] / 100

            # Criar uma cópia da alíquota para cálculos
            df['Alíquota_Calc'] = df['Alíquota (%)'].copy()

            # Verificação adicional para garantir que todos os valores estejam corretos
            print(f"Valores únicos de alíquota após correção: {df['Alíquota (%)'].unique()}")

        # Recalcular ISS Próprio para todas as linhas
        if 'ISS Próprio (R$)' in df.columns and 'Base de Cálculo (R$)' in df.columns and 'Alíquota_Calc' in df.columns and 'ISS Retido (R$)' in df.columns:
            # Calcular ISS Próprio com a alíquota ajustada
            df['ISS Próprio Calculado'] = df['Base de Cálculo (R$)'] * df['Alíquota_Calc'] / 100

            # Usar o valor calculado quando o valor original for zero ou muito diferente do calculado
            # MAS APENAS se o ISS Retido também for zero
            # Isso evita que o ISS Próprio seja recalculado quando o ISS Retido for diferente de zero
            mask = ((df['ISS Próprio (R$)'] == 0) & (df['ISS Retido (R$)'] == 0)) | \
                   ((abs(df['ISS Próprio (R$)'] - df['ISS Próprio Calculado']) > 0.01) & (df['ISS Retido (R$)'] == 0))
            df.loc[mask, 'ISS Próprio (R$)'] = df.loc[mask, 'ISS Próprio Calculado']

            # Remover as colunas temporárias
            df.drop(['ISS Próprio Calculado'], axis=1, inplace=True)

        # Remover a coluna de alíquota de cálculo após o uso
        if 'Alíquota_Calc' in df.columns:
            df.drop('Alíquota_Calc', axis=1, inplace=True)

        # Arredondar valores monetários para 2 casas decimais
        for col in numeric_columns:
            if col in df.columns:
                df[col] = df[col].round(2)

        # Arredondar alíquota para 2 casas decimais
        if 'Alíquota (%)' in df.columns:
            df['Alíquota (%)'] = df['Alíquota (%)'].round(2)

        return df

    except Exception as e:
        logger.error(f"Erro ao adicionar colunas calculadas: {str(e)}")
        return df

def filter_by_competence(df, start_month=None, end_month=None, year=None):
    """
    Filtra o DataFrame por competência (ano e intervalo de meses).

    Args:
        df (pd.DataFrame): DataFrame com os dados.
        start_month (int, optional): Mês inicial para filtrar (1-12).
        end_month (int, optional): Mês final para filtrar (1-12).
        year (int, optional): Ano para filtrar.

    Returns:
        pd.DataFrame: DataFrame filtrado.
    """
    try:
        # Verificar se o DataFrame tem as colunas necessárias
        required_columns = ['Competência', 'Mês', 'Ano']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            error_msg = f"Colunas necessárias ausentes no DataFrame: {missing_columns}"
            logger.error(error_msg)
            raise ValueError(error_msg)
            
        # Fazer uma cópia para não modificar o original
        filtered_df = df.copy()
        
        # Registrar informações sobre as competências antes da filtragem
        logger.info(f"Iniciando filtragem por competência: Ano={year}, Mês Inicial={start_month}, Mês Final={end_month}")
        logger.info(f"DataFrame original: {len(filtered_df)} linhas")
        logger.info(f"Competências únicas antes da filtragem: {filtered_df['Competência'].unique()}")
        logger.info(f"Valores únicos de mês antes da filtragem: {filtered_df['Mês'].unique()}")
        logger.info(f"Valores únicos de ano antes da filtragem: {filtered_df['Ano'].unique()}")
        
        # Verificar se há valores nulos nas colunas de mês e ano
        null_mes = filtered_df['Mês'].isna().sum()
        null_ano = filtered_df['Ano'].isna().sum()
        
        if null_mes > 0 or null_ano > 0:
            logger.warning(f"Detectados valores nulos: {null_mes} em 'Mês' e {null_ano} em 'Ano'")
            logger.warning("Tentando recuperar valores de mês e ano a partir da coluna 'Competência'")
            
            # Tentar recuperar valores de mês e ano para linhas com valores nulos
            null_mask = filtered_df['Mês'].isna() | filtered_df['Ano'].isna()
            competencias_problematicas = filtered_df.loc[null_mask, 'Competência'].unique()
            
            logger.info(f"Competências com valores nulos: {competencias_problematicas}")
            
            # Processar cada competência problemática
            for competencia in competencias_problematicas:
                if not isinstance(competencia, str) or competencia == 'nan':
                    continue
                    
                # Tentar extrair mês e ano usando regex mais flexível
                match = re.search(r'(\d{1,2})[^\d]+(\d{4})', competencia)
                if match:
                    try:
                        mes = int(match.group(1))
                        ano = int(match.group(2))
                        
                        # Validar mês
                        if 1 <= mes <= 12:
                            # Criar máscara para esta competência específica
                            comp_mask = (filtered_df['Competência'] == competencia) & null_mask
                            if comp_mask.any():
                                logger.info(f"Recuperando valores para {comp_mask.sum()} linhas com competência '{competencia}': mês={mes}, ano={ano}")
                                filtered_df.loc[comp_mask, 'Mês'] = float(mes)
                                filtered_df.loc[comp_mask, 'Ano'] = float(ano)
                        else:
                            logger.warning(f"Mês inválido ({mes}) encontrado na competência '{competencia}'")
                    except Exception as e:
                        logger.error(f"Erro ao processar competência '{competencia}': {str(e)}")
        
        # Verificar novamente valores nulos após a recuperação
        null_mes_after = filtered_df['Mês'].isna().sum()
        null_ano_after = filtered_df['Ano'].isna().sum()
        
        if null_mes_after > 0 or null_ano_after > 0:
            logger.warning(f"Após recuperação, ainda existem valores nulos: {null_mes_after} em 'Mês' e {null_ano_after} em 'Ano'")
            logger.warning("Estas linhas serão excluídas da filtragem por competência")
        
        # Aplicar filtro por intervalo de meses, se especificado
        if start_month is not None and end_month is not None:
            logger.info(f"Filtrando por intervalo de meses: de {start_month} a {end_month}")
            
            # Garante que os meses sejam tratados como números
            start_month = int(start_month)
            end_month = int(end_month)

            # Aplica o filtro de intervalo
            month_mask = (filtered_df['Mês'] >= start_month) & (filtered_df['Mês'] <= end_month)
            filtered_df = filtered_df[month_mask]
            
            logger.info(f"Após filtrar por intervalo de meses: {len(filtered_df)} linhas restantes")

        # Aplicar filtro por ano, se especificado
        if year is not None:
            if not isinstance(year, (int, float)):
                try:
                    year = int(year)
                except (ValueError, TypeError):
                    error_msg = f"Valor de ano inválido: {year}. Deve ser um número."
                    logger.error(error_msg)
                    raise ValueError(error_msg)
            
            logger.info(f"Filtrando por ano: {year}")
            # Verificar se há linhas com o ano especificado antes da filtragem
            year_count = (filtered_df['Ano'] == float(year)).sum()
            logger.info(f"Número de linhas com Ano={year} antes da filtragem: {year_count}")
            
            if year_count == 0:
                logger.warning(f"Nenhuma linha encontrada com Ano={year}")
                # Verificar se há competências que possam conter este ano
                potential_matches = []
                for comp in filtered_df['Competência'].unique():
                    if isinstance(comp, str) and f"/{year}" in comp:
                        potential_matches.append(comp)
                
                if potential_matches:
                    logger.warning(f"Possíveis competências para ano {year}: {potential_matches}")
            
            # Aplicar o filtro usando float para garantir compatibilidade de tipos
            filtered_df = filtered_df[filtered_df['Ano'] == float(year)]
            logger.info(f"Após filtrar por ano={year}: {len(filtered_df)} linhas restantes")
        
        # Registrar informações sobre as competências após a filtragem
        logger.info(f"Número de linhas após filtragem completa: {len(filtered_df)}")
        if not filtered_df.empty:
            logger.info(f"Competências únicas após filtragem: {filtered_df['Competência'].unique()}")
            logger.info(f"Valores únicos de mês após filtragem: {filtered_df['Mês'].unique()}")
            logger.info(f"Valores únicos de ano após filtragem: {filtered_df['Ano'].unique()}")
        else:
            logger.warning("Resultado da filtragem: DataFrame vazio")
            
            # Verificar se existem dados no DataFrame original que correspondam aos critérios
            if start_month is not None and end_month is not None and year is not None:
                # Procurar por competências que possam corresponder aos critérios
                # (Verifica qualquer mês dentro do intervalo)
                patterns = [f"{m:02d}/{year}" for m in range(start_month, end_month + 1)]
                matches = [comp for comp in df['Competência'].unique() if isinstance(comp, str) and any(p in comp for p in patterns)]
                matches = [comp for comp in df['Competência'].unique() if isinstance(comp, str) and pattern in comp]
                
                if matches:
                    logger.warning(f"Encontradas competências que correspondem aos critérios {pattern}: {matches}")
                    logger.warning("Tentando recuperar linhas com base na competência textual...")
                    
                    # Recuperar linhas com base na competência textual
                    recovered_rows = []
                    for match in matches:
                        match_mask = df['Competência'].str.contains(match, na=False)
                        if match_mask.any():
                            recovered_rows.append(df[match_mask].copy())
                    
                    if recovered_rows:
                        # Combinar as linhas recuperadas
                        recovered_df = pd.concat(recovered_rows)
                        # Forçar os valores de Mês e Ano para garantir que sejam filtrados corretamente
                        recovered_df['Mês'] = float(month)
                        recovered_df['Ano'] = float(year)
                        
                        filtered_df = recovered_df
                        logger.info(f"Recuperadas {len(filtered_df)} linhas com base na competência textual")
        
        return filtered_df
        
    except Exception as e:
        logger.error(f"Erro durante a filtragem por competência: {str(e)}")
        # Retornar o DataFrame original em caso de erro, para não perder dados
        logger.warning("Retornando DataFrame original devido a erro na filtragem")
        return df

def exclude_canceled_notes(df):
    """
    Exclui notas fiscais canceladas do DataFrame.

    Args:
        df (pd.DataFrame): DataFrame com os dados

    Returns:
        pd.DataFrame: DataFrame sem notas canceladas
    """
    # Assumindo que a coluna 'Situação' contém o status da nota
    # e que notas canceladas têm o valor 'CANCELADA'
    return df[df['Situação'] != 'CANCELADA']

def reorder_columns(df):
    """
    Reorganiza as colunas do DataFrame na ordem desejada.

    Args:
        df (pd.DataFrame): DataFrame com os dados

    Returns:
        pd.DataFrame: DataFrame com colunas reorganizadas
    """
    try:
        # Definir a ordem desejada das colunas
        # Após 'Data de Emissão' vem 'Competência'
        desired_order = [
            'Número da Nota',
            'Data de Emissão',
            'Competência',  # Movido para depois de Data de Emissão
            'CPF/CNPJ',
            'Nome do Tomador',
            'Código do Serviço',
            'Valor do Serviço (R$)',
            'Base de Cálculo (R$)',
            'Alíquota (%)',
            'ISS Próprio (R$)',
            'ISS Retido (R$)',
            'Natureza da Operação',
            'Incidência',
            'Situação',
            'Mês',
            'Ano'
        ]

        # Verificar quais colunas existem no DataFrame
        existing_columns = [col for col in desired_order if col in df.columns]

        # Adicionar quaisquer colunas que existam no DataFrame mas não estão na lista desejada
        for col in df.columns:
            if col not in existing_columns:
                existing_columns.append(col)

        # Reorganizar as colunas
        return df[existing_columns]

    except Exception as e:
        logger.error(f"Erro ao reorganizar colunas: {str(e)}")
        return df

def calculate_totals_by_competence(df):
    """
    Calcula totais por competência (mês/ano).

    Args:
        df (pd.DataFrame): DataFrame com os dados

    Returns:
        pd.DataFrame: DataFrame com totais por competência
    """
    # Agrupar por Mês e Ano e calcular totais
    totals = df.groupby(['Ano', 'Mês']).agg({
        'Valor do Serviço (R$)': 'sum',
        'Base de Cálculo (R$)': 'sum',
        'ISS Próprio (R$)': 'sum',
        'ISS Retido (R$)': 'sum',
        'Total ISS (R$)': 'sum',
        'Número da Nota': 'count'  # Conta o número de notas
    }).reset_index()

    # Renomear a coluna de contagem
    totals = totals.rename(columns={'Número da Nota': 'Quantidade de Notas'})

    return totals

def get_competence_list(df):
    """
    Obtém a lista de competências (mês/ano) presentes no DataFrame.

    Args:
        df (pd.DataFrame): DataFrame com os dados

    Returns:
        list: Lista de tuplas (ano, mês)
    """
    # Obter combinações únicas de Ano e Mês
    competences = df[['Ano', 'Mês']].drop_duplicates().sort_values(['Ano', 'Mês'])

    # Converter para lista de tuplas (ano, mês)
    competence_list = list(zip(competences['Ano'], competences['Mês']))

    return competence_list

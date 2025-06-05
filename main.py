import os
import sys
import argparse
import logging
from datetime import datetime

# Importar módulos do projeto
from pdf_extractor import extract_data_from_pdf
from data_processor import (
    process_data,
    filter_by_competence,
    exclude_canceled_notes,
    calculate_totals_by_competence
)
from excel_exporter import export_to_excel

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("pdf_to_excel.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("main")

def parse_arguments():
    """
    Analisa os argumentos da linha de comando.

    Returns:
        argparse.Namespace: Argumentos analisados
    """
    parser = argparse.ArgumentParser(
        description='Extrai dados de notas fiscais de um PDF e exporta para Excel.'
    )

    parser.add_argument(
        'pdf_path',
        type=str,
        help='Caminho para o arquivo PDF do Livro Anual de Serviços Prestados'
    )

    parser.add_argument(
        '--output',
        '-o',
        type=str,
        help='Caminho para salvar o arquivo Excel (opcional)'
    )

    parser.add_argument(
        '--month',
        '-m',
        type=int,
        choices=range(1, 13),
        help='Filtrar por mês (1-12)'
    )

    parser.add_argument(
        '--year',
        '-y',
        type=int,
        help='Filtrar por ano'
    )

    parser.add_argument(
        '--exclude-canceled',
        '-e',
        action='store_true',
        help='Excluir notas fiscais canceladas'
    )

    # Opção --no-separate removida, pois agora temos apenas duas planilhas fixas

    return parser.parse_args()

def process_pdf_file(pdf_path, output_dir, start_month=None, end_month=None, year=None, exclude_canceled=False):
    """
    Processa um único arquivo PDF e o exporta para Excel.

    Args:
        pdf_path (str): Caminho para o arquivo PDF.
        output_dir (str): Diretório onde o arquivo Excel será salvo.
        start_month (int, optional): Mês inicial para filtrar. Defaults to None.
        end_month (int, optional): Mês final para filtrar. Defaults to None.
        year (int, optional): Ano para filtrar. Defaults to None.
        exclude_canceled (bool, optional): Se True, exclui notas canceladas. Defaults to False.

    Returns:
        str: O caminho do arquivo Excel gerado ou None em caso de falha.
    """
    if not os.path.isfile(pdf_path):
        logger.error(f"Arquivo PDF não encontrado: {pdf_path}")
        return None

    try:
        logger.info(f"Processando arquivo PDF: {pdf_path}")
        df = extract_data_from_pdf(pdf_path)

        if df.empty:
            logger.warning("Nenhum dado extraído do PDF.")
            return None

        df = process_data(df)

        # O filtro deve ser aplicado se um ano for especificado, ou se o intervalo de meses for alterado.
        should_filter = year is not None or (start_month is not None and end_month is not None and (start_month != 1 or end_month != 12))
        
        if should_filter:
            log_year = year if year is not None else "Todos"
            logger.info(f"Aplicando filtro: Ano={log_year}, Meses de {start_month} a {end_month}")
            df = filter_by_competence(df, start_month, end_month, year)

        if exclude_canceled:
            logger.info("Excluindo notas fiscais canceladas")
            df = exclude_canceled_notes(df)

        if df.empty:
            logger.warning("Nenhum dado restante após aplicar filtros.")
            return None

        # Gera o nome do arquivo de saída
        pdf_basename = os.path.basename(pdf_path)
        pdf_name = os.path.splitext(pdf_basename)[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"{pdf_name}_{timestamp}.xlsx"
        output_path = os.path.join(output_dir, output_filename)

        os.makedirs(output_dir, exist_ok=True)

        excel_path = export_to_excel(df, output_path)

        if excel_path:
            logger.info(f"Arquivo Excel gerado com sucesso: {excel_path}")
            return excel_path
        else:
            logger.error("Falha ao gerar arquivo Excel.")
            return None

    except Exception as e:
        logger.error(f"Erro durante o processamento de {pdf_path}: {str(e)}")
        return None

def main():
    """
    Função principal do programa para execução via linha de comando.
    """
    args = parse_arguments()

    # A lógica da linha de comando agora suporta apenas um mês ou todos.
    # A GUI usará o intervalo completo.
    start_month_arg = args.month if args.month else 1
    end_month_arg = args.month if args.month else 12

    result_path = process_pdf_file(
        args.pdf_path,
        args.output or "output",
        start_month_arg,
        end_month_arg,
        args.year,
        args.exclude_canceled
    )

    if result_path:
        print(f"\nArquivo Excel gerado com sucesso: {result_path}")
    else:
        print("\nFalha ao processar o arquivo.")
        sys.exit(1)

if __name__ == "__main__":
    print("\n=== Extrator de Notas Fiscais PDF para Excel ===\n")
    main()
    print("\nProcessamento concluído.\n")

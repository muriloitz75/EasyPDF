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

def main():
    """
    Função principal do programa.
    """
    # Analisar argumentos da linha de comando
    args = parse_arguments()

    # Verificar se o arquivo PDF existe
    if not os.path.isfile(args.pdf_path):
        logger.error(f"Arquivo PDF não encontrado: {args.pdf_path}")
        sys.exit(1)

    try:
        # Extrair dados do PDF
        logger.info(f"Processando arquivo PDF: {args.pdf_path}")
        df = extract_data_from_pdf(args.pdf_path)

        if df.empty:
            logger.error("Nenhum dado extraído do PDF.")
            sys.exit(1)

        # Processar dados
        df = process_data(df)

        # Aplicar filtros, se solicitado
        if args.month or args.year:
            logger.info(f"Aplicando filtro: Mês={args.month}, Ano={args.year}")
            df = filter_by_competence(df, args.month, args.year)

        # Excluir notas canceladas, se solicitado
        if args.exclude_canceled:
            logger.info("Excluindo notas fiscais canceladas")
            df = exclude_canceled_notes(df)

        # Verificar se ainda há dados após os filtros
        if df.empty:
            logger.error("Nenhum dado restante após aplicar filtros.")
            sys.exit(1)

        # Definir caminho de saída, se não fornecido
        output_path = args.output
        if output_path is None:
            # Usar a pasta output como diretório padrão
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join("output", f"notas_fiscais_{timestamp}.xlsx")

            # Garantir que a pasta output existe
            os.makedirs("output", exist_ok=True)

        # Exportar para Excel
        excel_path = export_to_excel(df, output_path)

        if excel_path:
            logger.info(f"Arquivo Excel gerado com sucesso: {excel_path}")
            print(f"\nArquivo Excel gerado com sucesso: {excel_path}")
        else:
            logger.error("Falha ao gerar arquivo Excel.")
            sys.exit(1)

    except Exception as e:
        logger.error(f"Erro durante o processamento: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    print("\n=== Extrator de Notas Fiscais PDF para Excel ===\n")
    main()
    print("\nProcessamento concluído.\n")

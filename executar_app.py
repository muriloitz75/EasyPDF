"""
Script para executar a aplicação e manter a janela aberta.
"""
import os
import sys
import subprocess
import glob
from datetime import datetime

def main():
    """
    Função principal que executa a aplicação.
    """
    print("\n===== EXTRATOR DE NOTAS FISCAIS PDF PARA EXCEL =====\n")

    # Criar pastas input e output se não existirem
    if not os.path.exists("input"):
        os.makedirs("input")
        print("Pasta 'input' criada.")

    if not os.path.exists("output"):
        os.makedirs("output")
        print("Pasta 'output' criada.")

    # Procurar arquivos PDF na pasta input
    pdf_files = []

    # Verificar se foi fornecido um arquivo PDF como argumento
    if len(sys.argv) > 1:
        pdf_file = sys.argv[1]
        if os.path.exists(pdf_file) and pdf_file.lower().endswith('.pdf'):
            pdf_files.append(pdf_file)
        else:
            print(f"Arquivo não encontrado ou não é um PDF: {pdf_file}")
    else:
        # Procurar arquivos PDF na pasta input
        pdf_files = glob.glob("input/*.pdf")

    if not pdf_files:
        print("Nenhum arquivo PDF encontrado na pasta 'input'.")
        print("Por favor, coloque seus arquivos PDF na pasta 'input' e execute novamente.")
        print("\nExemplo de uso:")
        print("  python executar_app.py caminho/para/arquivo.pdf")
        print("  ou coloque arquivos PDF na pasta 'input' e execute sem argumentos")
    else:
        print(f"Encontrados {len(pdf_files)} arquivos PDF para processar.")
        for pdf_file in pdf_files:
            process_pdf(pdf_file)

    print("\nProcessamento concluído.")
    print("\nPressione ENTER para sair...")
    input()  # Mantém a janela aberta até o usuário pressionar ENTER

def process_pdf(pdf_file):
    """
    Processa um arquivo PDF usando a aplicação principal.

    Args:
        pdf_file (str): Caminho para o arquivo PDF
    """
    print(f"Processando arquivo: {pdf_file}")

    try:
        # Gerar nome de saída baseado no nome do arquivo PDF
        pdf_basename = os.path.basename(pdf_file)
        pdf_name = os.path.splitext(pdf_basename)[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join("output", f"{pdf_name}_{timestamp}.xlsx")

        # Executar a aplicação principal
        cmd = [sys.executable, "main.py", pdf_file, "-o", output_file]
        print(f"Executando comando: {' '.join(cmd)}")

        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode == 0:
            print(f"Arquivo processado com sucesso: {output_file}")
            if result.stdout:
                print(result.stdout)
        else:
            print(f"Erro ao processar o arquivo. Código de saída: {result.returncode}")
            if result.stdout:
                print("Saída padrão:")
                print(result.stdout)
            if result.stderr:
                print("Saída de erro:")
                print(result.stderr)

    except Exception as e:
        print(f"Erro ao processar o arquivo: {str(e)}")

    print()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Erro inesperado: {str(e)}")
        print("\nPressione ENTER para sair...")
        input()  # Mantém a janela aberta mesmo em caso de erro

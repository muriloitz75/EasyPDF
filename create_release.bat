@echo off
echo "Criando pacote de distribuicao..."
if exist EasyPDF-Extractor-v1.0.zip del EasyPDF-Extractor-v1.0.zip
tar -a -c -f EasyPDF-Extractor-v1.0.zip dist/EasyPDF-Extractor-GUI-Final.exe README.md INSTRUCOES_EXECUCAO.md
echo "Pacote de distribuicao criado com sucesso: EasyPDF-Extractor-v1.0.zip"
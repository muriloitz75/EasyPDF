# Solução para o Executável que Fecha Imediatamente

Se o executável `PDFtoExcel.exe` estiver fechando imediatamente após ser executado, siga uma das soluções abaixo:

## Solução 1: Usar o Arquivo de Lote

1. Execute o arquivo `Executar_Aplicacao.bat` que foi criado para resolver este problema
2. Este arquivo irá:
   - Verificar se o Python está instalado
   - Instalar as dependências necessárias
   - Executar a aplicação
   - Manter a janela aberta até você pressionar uma tecla

## Solução 2: Executar via Prompt de Comando (CMD)

1. Abra o Prompt de Comando (CMD) como administrador
2. Navegue até a pasta onde você extraiu o arquivo ZIP:
   ```
   cd caminho\para\pasta\extraida
   ```
3. Execute o comando:
   ```
   PDFtoExcel\PDFtoExcel.exe input\nome_do_arquivo.pdf
   ```

## Solução 3: Executar como Administrador

1. Clique com o botão direito no arquivo `PDFtoExcel\PDFtoExcel.exe`
2. Selecione "Executar como administrador"

## Solução 4: Usar o Python Diretamente

Se você tiver o Python instalado, pode executar a aplicação diretamente:

1. Abra o Prompt de Comando (CMD)
2. Navegue até a pasta onde você extraiu o arquivo ZIP
3. Execute o comando:
   ```
   python main.py input\nome_do_arquivo.pdf
   ```

## Requisitos do Sistema

- Windows 7 ou superior
- Python 3.8 ou superior (para as soluções 1 e 4)
- Permissões de administrador podem ser necessárias
- Espaço em disco: pelo menos 100 MB livres

## Solução de Problemas Comuns

1. **Erro "Arquivo não encontrado"**:
   - Verifique se o arquivo PDF está na pasta `input`
   - Verifique se o caminho está correto

2. **Erro "Acesso negado"**:
   - Execute o programa como administrador
   - Verifique as permissões da pasta

3. **Programa não inicia**:
   - Verifique se o Windows Defender ou antivírus está bloqueando a execução
   - Tente adicionar uma exceção no antivírus

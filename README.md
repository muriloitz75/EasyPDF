# Extrator de Notas Fiscais PDF para Excel

Aplicação para extrair dados de notas fiscais do Livro Anual de Serviços Prestados e exportar para Excel.

## Melhorias Implementadas

### 1. Sistema Robusto de Reconhecimento de Padrões de Competência

Implementamos um sistema avançado para reconhecer diferentes formatos de competência nos PDFs:

- Reconhecimento de múltiplos formatos: MM/AAAA, Competência: MM/AAAA, MM-AAAA, MM.AAAA
- Processamento individual de cada competência para maior precisão
- Validação de mês e ano para evitar valores incorretos
- Recuperação inteligente de valores quando os padrões padrão falham

### 2. Tratamento de Erros Aprimorado

- Mensagens de erro mais específicas e detalhadas
- Logs abrangentes para facilitar a depuração
- Recuperação de falhas para evitar interrupções no processamento
- Validação de dados de entrada para prevenir erros

### 3. Extração de PDF Mais Robusta

- Verificação de existência e validade do arquivo PDF
- Tratamento de diferentes layouts e formatos de página
- Recuperação de dados mesmo quando as tabelas não são reconhecidas
- Processamento página a página com tratamento de erros independente

### 4. Testes Automatizados

Adicionamos testes automatizados para garantir a qualidade do código:

- Testes para o processador de dados (`test_data_processor.py`)
- Testes para o extrator de PDF (`test_pdf_extractor.py`)
- Verificação de diferentes formatos de competência
- Testes de robustez com valores nulos e inválidos

#### Como Executar os Testes

Para executar os testes automatizados, siga estas instruções:

```
# Executar testes específicos
python test_data_processor.py
python test_pdf_extractor.py
```

Os logs dos testes serão salvos em `test_data_processor.log` e `test_pdf_extractor.log` para análise detalhada.

## SOLUÇÃO DE PROBLEMAS COM O EXECUTÁVEL

Se o executável não estiver funcionando corretamente, siga estas instruções:

### Método 1: Executar via Prompt de Comando (CMD)

1. Abra o Prompt de Comando (CMD) como administrador
2. Navegue até a pasta onde você extraiu o arquivo ZIP usando o comando `cd`
   ```
   cd caminho\para\pasta\extraida
   ```
3. Execute o comando:
   ```
   PDFtoExcel\PDFtoExcel.exe input\nome_do_arquivo.pdf
   ```

### Método 2: Verificar Permissões

1. Clique com o botão direito na pasta `PDFtoExcel`
2. Selecione "Propriedades"
3. Vá para a aba "Segurança"
4. Clique em "Editar" e depois em "Adicionar"
5. Adicione permissões completas para o seu usuário
6. Clique em "Aplicar" e depois em "OK"

### Método 3: Executar como Administrador

1. Clique com o botão direito no arquivo `PDFtoExcel\PDFtoExcel.exe`
2. Selecione "Executar como administrador"

## INSTRUÇÕES BÁSICAS DE USO

1. **Coloque os arquivos PDF** na pasta `input`
2. **Execute o programa** usando um dos métodos acima
3. **Encontre os arquivos Excel** gerados na pasta `output`

## Funcionalidades

- Extração de dados de notas fiscais de arquivos PDF
- Processamento e estruturação dos dados extraídos
- Filtragem por competência (mês/ano)
- Exclusão de notas fiscais canceladas
- Exportação para Excel com duas planilhas: "Dados_NFSe" e "Resumo_NFSe"
- Resumo com totais por competência

## Estrutura de Pastas

- **input**: Coloque aqui os arquivos PDF a serem processados
- **output**: Os arquivos Excel gerados serão salvos nesta pasta

### Opções Disponíveis

- `--output`, `-o`: Caminho para salvar o arquivo Excel (opcional)
- `--month`, `-m`: Filtrar por mês (1-12)
- `--year`, `-y`: Filtrar por ano
- `--exclude-canceled`, `-e`: Excluir notas fiscais canceladas

### Exemplos de Comandos

Processar um arquivo PDF específico:
```
PDFtoExcel\PDFtoExcel.exe input\arquivo.pdf
```

Extrair notas fiscais de um mês específico:
```
PDFtoExcel\PDFtoExcel.exe input\arquivo.pdf -m 3 -y 2025
```

Extrair notas fiscais e excluir canceladas:
```
PDFtoExcel\PDFtoExcel.exe input\arquivo.pdf -e
```

Especificar arquivo de saída:
```
PDFtoExcel\PDFtoExcel.exe input\arquivo.pdf -o output\saida.xlsx
```

## Requisitos do Sistema

- Windows 7 ou superior
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

## Campos Extraídos

- Número da nota
- Data de emissão
- CPF/CNPJ do tomador
- Nome do tomador do serviço
- Código do serviço
- Valor do serviço (R$)
- Base de cálculo (R$)
- Alíquota (%)
- ISS próprio (R$)
- ISS retido (R$)
- Natureza da operação
- Incidência
- Situação (ESCRITURADA, CANCELADA, etc.)
- Competência (mês/ano)

## Observações

- O padrão de extração de texto pode precisar de ajustes dependendo do formato exato do PDF
- Recomenda-se verificar os primeiros resultados para garantir a precisão da extração
- Os logs são salvos em arquivos `.log` para facilitar a depuração

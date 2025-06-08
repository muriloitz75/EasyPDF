import unittest
import pandas as pd
import numpy as np
import os
import sys
import logging
from datetime import datetime

# Adicionar o diretório pai ao path para importar os módulos
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Importar os módulos a serem testados
from data_processor import (
    process_data,
    add_calculated_columns,
    filter_by_competence,
    clean_tomador_name
)

# Configurar logging para testes
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("test_data_processor.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("test_data_processor")

class TestDataProcessor(unittest.TestCase):
    """Testes para o módulo data_processor"""
    
    def setUp(self):
        """Configuração inicial para os testes"""
        # Criar um DataFrame de teste com diferentes formatos de competência
        self.test_data = pd.DataFrame({
            'Número da Nota': ['123456', '234567', '345678', '456789', '567890'],
            'Data de Emissão': ['01/01/2023', '02/02/2023', '03/03/2023', '04/04/2023', '05/05/2023'],
            'Competência': ['01/2023', 'Competência: 02/2023', 'COMPETÊNCIA:03/2023', '04-2023', '05.2023'],
            'Nome do Tomador': ['EMPRESA A', '50 BATALHÃO DE INFANTARIA', 'EMPRESA C 0403', 'EMPRESA D 50', 'EMPRESA E 100,50'],
            'CPF/CNPJ': ['12345678901234', '23456789012345', '34567890123456', '45678901234567', '56789012345678'],
            'Valor do Serviço (R$)': [1000.0, 2000.0, 3000.0, 4000.0, 5000.0],
            'Base de Cálculo (R$)': [1000.0, 3.0, 3000.0, 4000.0, 5000.0],
            'Alíquota (%)': [5.0, 0.0, 300.0, 2.0, 3.0],
            'ISS Próprio (R$)': [50.0, 0.0, 90.0, 80.0, 150.0],
            'ISS Retido (R$)': [0.0, 0.0, 0.0, 0.0, 0.0],
            'Situação': ['NORMAL', 'NORMAL', 'NORMAL', 'CANCELADA', 'NORMAL']
        })
        
        # Criar um DataFrame com valores nulos para testar robustez
        self.null_data = pd.DataFrame({
            'Número da Nota': ['123456', '234567'],
            'Data de Emissão': ['01/01/2023', '02/02/2023'],
            'Competência': ['01/2023', None],
            'Nome do Tomador': ['EMPRESA A', None],
            'CPF/CNPJ': ['12345678901234', '23456789012345'],
            'Valor do Serviço (R$)': [1000.0, 2000.0],
            'Base de Cálculo (R$)': [1000.0, 2000.0],
            'Alíquota (%)': [5.0, 3.0],
            'ISS Próprio (R$)': [50.0, 60.0],
            'ISS Retido (R$)': [0.0, 0.0],
            'Situação': ['NORMAL', 'NORMAL']
        })
    
    def test_process_data(self):
        """Testar o processamento completo de dados"""
        processed_df = process_data(self.test_data)
        
        # Verificar se o DataFrame foi processado corretamente
        self.assertFalse(processed_df.empty)
        self.assertEqual(len(processed_df), len(self.test_data))
        
        # Verificar se as colunas Mês e Ano foram adicionadas
        self.assertIn('Mês', processed_df.columns)
        self.assertIn('Ano', processed_df.columns)
        
        # Verificar se os tipos de dados foram convertidos corretamente
        self.assertTrue(pd.api.types.is_datetime64_any_dtype(processed_df['Data de Emissão']))
        self.assertTrue(pd.api.types.is_numeric_dtype(processed_df['Valor do Serviço (R$)']))
        
        # Verificar se o nome do tomador foi limpo
        self.assertNotIn('0403', processed_df.loc[2, 'Nome do Tomador'])
        self.assertNotIn('50', processed_df.loc[3, 'Nome do Tomador'])
    
    def test_add_calculated_columns(self):
        """Testar a adição de colunas calculadas"""
        # Processar o DataFrame de teste usando process_data que chama convert_data_types e add_calculated_columns
        df_copy = self.test_data.copy()
        result_df = process_data(df_copy)
        
        # Verificar se as colunas Mês e Ano foram adicionadas
        self.assertIn('Mês', result_df.columns)
        self.assertIn('Ano', result_df.columns)
        
        # Verificar se os valores de Mês e Ano foram extraídos corretamente
        # para diferentes formatos de competência
        self.assertEqual(result_df.loc[0, 'Mês'], 1.0)  # 01/2023
        self.assertEqual(result_df.loc[0, 'Ano'], 2023.0)
        
        self.assertEqual(result_df.loc[1, 'Mês'], 2.0)  # Competência: 02/2023
        self.assertEqual(result_df.loc[1, 'Ano'], 2023.0)
        
        self.assertEqual(result_df.loc[2, 'Mês'], 3.0)  # COMPETÊNCIA:03/2023
        self.assertEqual(result_df.loc[2, 'Ano'], 2023.0)
        
        self.assertEqual(result_df.loc[3, 'Mês'], 4.0)  # 04-2023
        self.assertEqual(result_df.loc[3, 'Ano'], 2023.0)
        
        self.assertEqual(result_df.loc[4, 'Mês'], 5.0)  # 05.2023
        self.assertEqual(result_df.loc[4, 'Ano'], 2023.0)
        
        # Verificar se a alíquota foi corrigida (de 300 para 3)
        self.assertEqual(result_df.loc[2, 'Alíquota (%)'], 3.0)
        
        # Verificar se a Alíquota foi corrigida para a linha com BATALHÃO DE INFANTARIA
        # A Base de Cálculo só é atualizada quando a Base de Cálculo é 3.0 e a Alíquota é 0.0
        self.assertEqual(result_df.loc[1, 'Alíquota (%)'], 3.0)  # Deve ser corrigida para 3.0
    
    def test_filter_by_competence(self):
        """Testar a filtragem por competência"""
        # Processar o DataFrame de teste primeiro
        df_copy = self.test_data.copy()
        processed_df = process_data(df_copy)
        
        # Testar filtragem por mês
        filtered_by_month = filter_by_competence(processed_df, start_month=1, end_month=1)
        self.assertEqual(len(filtered_by_month), 1)
        self.assertEqual(filtered_by_month.iloc[0]['Mês'], 1.0)
        
        # Testar filtragem por ano
        filtered_by_year = filter_by_competence(processed_df, year=2023)
        self.assertEqual(len(filtered_by_year), 5)  # Todos os registros são de 2023
        
        # Testar filtragem por mês e ano
        filtered_by_both = filter_by_competence(processed_df, start_month=2, end_month=2, year=2023)
        self.assertEqual(len(filtered_by_both), 1)
        self.assertEqual(filtered_by_both.iloc[0]['Mês'], 2.0)
        self.assertEqual(filtered_by_both.iloc[0]['Ano'], 2023.0)
        
        # Testar filtragem com valores que não existem
        filtered_empty = filter_by_competence(processed_df, start_month=6, end_month=6, year=2024)
        self.assertTrue(filtered_empty.empty)
    
    def test_clean_tomador_name(self):
        """Testar a limpeza do nome do tomador"""
        # Testar casos específicos
        self.assertEqual(clean_tomador_name('EMPRESA A 0403'), 'EMPRESA A')
        self.assertEqual(clean_tomador_name('50 BATALHÃO DE INFANTARIA'), 'BATALHÃO DE INFANTARIA')
        self.assertEqual(clean_tomador_name('EMPRESA C 100,50'), 'EMPRESA C')
        
        # Testar com valor nulo
        self.assertTrue(pd.isna(clean_tomador_name(np.nan)))
    
    def test_robustness_with_null_values(self):
        """Testar a robustez do processamento com valores nulos"""
        # Processar DataFrame com valores nulos
        processed_df = process_data(self.null_data)
        
        # Verificar se o processamento não falhou com valores nulos
        self.assertFalse(processed_df.empty)
        self.assertEqual(len(processed_df), len(self.null_data))
        
        # Verificar se as linhas com competência nula foram tratadas
        # A linha com competência nula deve ter Mês e Ano como NaN
        self.assertTrue(pd.isna(processed_df.loc[1, 'Mês']))
        self.assertTrue(pd.isna(processed_df.loc[1, 'Ano']))

if __name__ == '__main__':
    unittest.main()
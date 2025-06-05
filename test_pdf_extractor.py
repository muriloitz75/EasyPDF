import unittest
import pandas as pd
import os
import sys
import logging
import tempfile
from unittest.mock import patch, MagicMock

# Adicionar o diretório pai ao path para importar os módulos
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Importar os módulos a serem testados
from pdf_extractor import (
    extract_data_from_pdf,
    extract_from_tables,
    extract_notes_from_text
)

# Configurar logging para testes
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("test_pdf_extractor.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("test_pdf_extractor")

class TestPdfExtractor(unittest.TestCase):
    """Testes para o módulo pdf_extractor"""
    
    def test_extract_notes_from_text(self):
        """Testar a extração de notas fiscais a partir de texto"""
        # Texto de exemplo simulando uma nota fiscal no formato esperado pela função
        # O formato deve corresponder ao padrão definido em note_pattern na função extract_notes_from_text
        sample_text = """COMPETÊNCIA: 01/2023
        
        123456789012345 01/01/2023 12345678901234 EMPRESA A 0123 1000,00 1000,00 5,00 % 50,00 0,00 EXIGÍVEL ESCRITURADA
        234567890123456 02/02/2023 23456789012345 EMPRESA B 0456 2000,00 2000,00 3,00 % 60,00 0,00 EXIGÍVEL ESCRITURADA
        """
        
        # Extrair notas do texto
        notes = extract_notes_from_text(sample_text)
        
        # Verificar se as notas foram extraídas corretamente
        self.assertEqual(len(notes), 2)
        self.assertEqual(notes[0]['Número da Nota'], '123456789012345')
        self.assertEqual(notes[1]['Número da Nota'], '234567890123456')
    
    def test_extract_data_from_pdf_with_different_competence_formats(self):
        """Testar a extração de dados de PDF com diferentes formatos de competência"""
        # Testar diretamente a função extract_notes_from_text com diferentes formatos de competência
        
        # Formato 1: COMPETÊNCIA: MM/AAAA
        text1 = "COMPETÊNCIA: 01/2023\n123456789012345 01/01/2023 12345678901234 EMPRESA A 0123 1000,00 1000,00 5,00 % 50,00 0,00 EXIGÍVEL ESCRITURADA"
        notes1 = extract_notes_from_text(text1)
        self.assertEqual(len(notes1), 1)
        self.assertEqual(notes1[0]['Competência'], '01/2023')
        
        # Formato 2: Competência MM/AAAA
        text2 = "Competência 02/2023\n234567890123456 02/02/2023 23456789012345 EMPRESA B 0456 2000,00 2000,00 3,00 % 60,00 0,00 EXIGÍVEL ESCRITURADA"
        notes2 = extract_notes_from_text(text2)
        self.assertEqual(len(notes2), 1)
        self.assertEqual(notes2[0]['Competência'], '02/2023')
        
        # Formato 3: MM/AAAA (formato de data)
        text3 = "Relatório de Notas Fiscais 03/2023\n345678901234567 03/03/2023 34567890123456 EMPRESA C 0789 3000,00 3000,00 2,00 % 60,00 0,00 EXIGÍVEL ESCRITURADA"
        notes3 = extract_notes_from_text(text3)
        self.assertEqual(len(notes3), 1)
        self.assertEqual(notes3[0]['Competência'], '03/2023')
    
    def test_extract_data_from_pdf_error_handling(self):
        """Testar o tratamento de erros na extração de dados de PDF"""
        # Testar com um arquivo que não existe
        with self.assertRaises(FileNotFoundError):
            extract_data_from_pdf('arquivo_inexistente.pdf')
        
        # Testar com um arquivo que não é um PDF válido
        with tempfile.NamedTemporaryFile(suffix='.txt') as temp_file:
            temp_file.write(b'Este arquivo nao e um PDF valido')
            temp_file.flush()
            
            # Deve lançar uma exceção ou retornar um DataFrame vazio
            with self.assertRaises(ValueError):
                extract_data_from_pdf(temp_file.name)

if __name__ == '__main__':
    unittest.main()
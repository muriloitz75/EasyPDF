#!/usr/bin/env python3
"""
Launcher para a interface gráfica do EasyPDF Extractor
"""

import sys
import os

# Adicionar o diretório atual ao path para importações
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    import customtkinter as ctk
    from gui import App
    
    if __name__ == "__main__":
        print("Iniciando EasyPDF Extractor GUI...")
        app = App()
        app.mainloop()
        
except ImportError as e:
    print(f"Erro de importação: {e}")
    print("Verifique se todas as dependências estão instaladas.")
    input("Pressione Enter para sair...")
except Exception as e:
    print(f"Erro inesperado: {e}")
    input("Pressione Enter para sair...")
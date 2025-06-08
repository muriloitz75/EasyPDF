#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de teste para diagnosticar problemas com a GUI do EasyPDF Extractor
"""

import sys
import traceback
import os

print("=== TESTE DE DIAGNÓSTICO DA GUI ===")
print(f"Python: {sys.version}")
print(f"Diretório atual: {os.getcwd()}")
print()

try:
    print("1. Testando importação do tkinter...")
    import tkinter as tk
    print("   ✓ tkinter importado com sucesso")
except Exception as e:
    print(f"   ✗ Erro ao importar tkinter: {e}")
    sys.exit(1)

try:
    print("2. Testando importação do customtkinter...")
    import customtkinter as ctk
    print("   ✓ customtkinter importado com sucesso")
except Exception as e:
    print(f"   ✗ Erro ao importar customtkinter: {e}")
    sys.exit(1)

try:
    print("3. Testando criação de janela básica...")
    root = ctk.CTk()
    root.title("Teste")
    root.geometry("300x200")
    print("   ✓ Janela criada com sucesso")
    
    # Teste rápido sem mainloop
    root.update()
    print("   ✓ Update da janela funcionou")
    
    root.destroy()
    print("   ✓ Janela destruída com sucesso")
except Exception as e:
    print(f"   ✗ Erro ao criar janela: {e}")
    traceback.print_exc()
    sys.exit(1)

try:
    print("4. Testando importação dos módulos do projeto...")
    from main import process_pdf_file
    print("   ✓ Módulo main importado com sucesso")
except Exception as e:
    print(f"   ✗ Erro ao importar módulo main: {e}")
    traceback.print_exc()

try:
    print("5. Testando importação completa da GUI...")
    import gui
    print("   ✓ Módulo gui importado com sucesso")
except Exception as e:
    print(f"   ✗ Erro ao importar módulo gui: {e}")
    traceback.print_exc()
    sys.exit(1)

try:
    print("6. Testando criação da classe App...")
    app = gui.App()
    print("   ✓ Classe App criada com sucesso")
    
    # Teste sem mainloop para evitar travamento
    app.update()
    print("   ✓ Update da App funcionou")
    
    app.destroy()
    print("   ✓ App destruída com sucesso")
except Exception as e:
    print(f"   ✗ Erro ao criar App: {e}")
    traceback.print_exc()
    sys.exit(1)

print()
print("=== TODOS OS TESTES PASSARAM ===")
print("A GUI deveria funcionar normalmente.")
print("Se o executável não abre a GUI, pode ser um problema de compilação.")
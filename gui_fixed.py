#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Versão corrigida da GUI do EasyPDF Extractor
Corrige problemas de threading e execução em executáveis
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import glob
import threading
import sys
from datetime import datetime

# Importar função de processamento
try:
    from main import process_pdf_file
except ImportError:
    def process_pdf_file(*args, **kwargs):
        raise ImportError("Módulo main não encontrado")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configurar aparência
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        
        # Configurar janela principal
        self.title("EasyPDF Extractor - Extração de Notas Fiscais")
        self.geometry("800x600")
        self.minsize(600, 400)
        
        # Centralizar janela
        self.center_window()
        
        # Variáveis de estado
        self.input_folder = tk.StringVar(value="")
        self.output_folder = tk.StringVar(value="")
        self.start_month = tk.StringVar(value="1")
        self.end_month = tk.StringVar(value="12")
        self.processing = False
        
        # Criar interface
        self.create_widgets()
        
        # Configurar protocolo de fechamento
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def center_window(self):
        """Centraliza a janela na tela"""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
    
    def create_widgets(self):
        """Cria todos os widgets da interface"""
        # Frame principal
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Título
        title_label = ctk.CTkLabel(
            main_frame, 
            text="EasyPDF Extractor", 
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=(20, 30))
        
        # Frame de configurações
        config_frame = ctk.CTkFrame(main_frame)
        config_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        # Pasta de entrada
        input_frame = ctk.CTkFrame(config_frame)
        input_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(input_frame, text="Pasta de Entrada (PDFs):").pack(anchor="w", padx=10, pady=(10, 5))
        
        input_path_frame = ctk.CTkFrame(input_frame)
        input_path_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        self.input_entry = ctk.CTkEntry(input_path_frame, textvariable=self.input_folder)
        self.input_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        input_button = ctk.CTkButton(
            input_path_frame, 
            text="Selecionar", 
            command=self.select_input_folder,
            width=100
        )
        input_button.pack(side="right")
        
        # Pasta de saída
        output_frame = ctk.CTkFrame(config_frame)
        output_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(output_frame, text="Pasta de Saída (Excel):").pack(anchor="w", padx=10, pady=(10, 5))
        
        output_path_frame = ctk.CTkFrame(output_frame)
        output_path_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        self.output_entry = ctk.CTkEntry(output_path_frame, textvariable=self.output_folder)
        self.output_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        output_button = ctk.CTkButton(
            output_path_frame, 
            text="Selecionar", 
            command=self.select_output_folder,
            width=100
        )
        output_button.pack(side="right")
        
        # Filtros de mês
        filter_frame = ctk.CTkFrame(config_frame)
        filter_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(filter_frame, text="Filtro por Mês:").pack(anchor="w", padx=10, pady=(10, 5))
        
        month_frame = ctk.CTkFrame(filter_frame)
        month_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        ctk.CTkLabel(month_frame, text="De:").pack(side="left", padx=(10, 5))
        
        self.start_month_combo = ctk.CTkComboBox(
            month_frame,
            values=[str(i) for i in range(1, 13)],
            variable=self.start_month,
            width=80
        )
        self.start_month_combo.pack(side="left", padx=5)
        
        ctk.CTkLabel(month_frame, text="Até:").pack(side="left", padx=(20, 5))
        
        self.end_month_combo = ctk.CTkComboBox(
            month_frame,
            values=[str(i) for i in range(1, 13)],
            variable=self.end_month,
            width=80
        )
        self.end_month_combo.pack(side="left", padx=5)
        
        # Botão de processamento
        self.process_button = ctk.CTkButton(
            main_frame,
            text="Processar PDFs",
            command=self.start_processing,
            height=40,
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.process_button.pack(pady=20)
        
        # Barra de progresso
        self.progress_bar = ctk.CTkProgressBar(main_frame)
        self.progress_bar.pack(fill="x", padx=20, pady=(0, 10))
        self.progress_bar.set(0)
        
        # Log de saída
        log_frame = ctk.CTkFrame(main_frame)
        log_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        ctk.CTkLabel(log_frame, text="Log de Processamento:").pack(anchor="w", padx=10, pady=(10, 5))
        
        self.log_textbox = ctk.CTkTextbox(log_frame, height=150)
        self.log_textbox.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Log inicial
        self.log_message("EasyPDF Extractor iniciado. Selecione as pastas e clique em 'Processar PDFs'.")
    
    def select_input_folder(self):
        """Seleciona pasta de entrada"""
        folder = filedialog.askdirectory(title="Selecionar pasta com arquivos PDF")
        if folder:
            self.input_folder.set(folder)
            self.log_message(f"Pasta de entrada selecionada: {folder}")
    
    def select_output_folder(self):
        """Seleciona pasta de saída"""
        folder = filedialog.askdirectory(title="Selecionar pasta para salvar arquivos Excel")
        if folder:
            self.output_folder.set(folder)
            self.log_message(f"Pasta de saída selecionada: {folder}")
    
    def start_processing(self):
        """Inicia o processamento em thread separada"""
        if self.processing:
            return
        
        # Validar entradas
        if not self.input_folder.get():
            messagebox.showerror("Erro", "Selecione a pasta de entrada")
            return
        
        if not self.output_folder.get():
            messagebox.showerror("Erro", "Selecione a pasta de saída")
            return
        
        # Verificar se há arquivos PDF
        pdf_files = glob.glob(os.path.join(self.input_folder.get(), "*.pdf"))
        if not pdf_files:
            messagebox.showerror("Erro", "Nenhum arquivo PDF encontrado na pasta de entrada")
            return
        
        # Iniciar processamento
        self.processing = True
        self.process_button.configure(text="Processando...", state="disabled")
        self.progress_bar.set(0)
        
        # Executar em thread separada
        thread = threading.Thread(target=self.process_files, daemon=True)
        thread.start()
    
    def process_files(self):
        """Processa os arquivos PDF"""
        try:
            input_dir = self.input_folder.get()
            output_dir = self.output_folder.get()
            start_month = int(self.start_month.get())
            end_month = int(self.end_month.get())
            
            self.log_message(f"Iniciando processamento...")
            self.log_message(f"Pasta de entrada: {input_dir}")
            self.log_message(f"Pasta de saída: {output_dir}")
            self.log_message(f"Filtro de mês: {start_month} a {end_month}")
            
            # Encontrar arquivos PDF
            pdf_files = glob.glob(os.path.join(input_dir, "*.pdf"))
            self.log_message(f"Encontrados {len(pdf_files)} arquivo(s) PDF")
            
            # Processar cada arquivo
            for i, pdf_file in enumerate(pdf_files):
                try:
                    self.log_message(f"Processando: {os.path.basename(pdf_file)}")
                    
                    # Atualizar progresso
                    progress = (i + 1) / len(pdf_files)
                    self.after(0, lambda p=progress: self.progress_bar.set(p))
                    
                    # Processar arquivo
                    result = process_pdf_file(
                        pdf_file, 
                        output_dir, 
                        start_month, 
                        end_month
                    )
                    
                    if result:
                        self.log_message(f"✓ Processado com sucesso: {os.path.basename(result)}")
                    else:
                        self.log_message(f"✗ Erro ao processar: {os.path.basename(pdf_file)}")
                        
                except Exception as e:
                    self.log_message(f"✗ Erro ao processar {os.path.basename(pdf_file)}: {str(e)}")
            
            self.log_message("\n=== PROCESSAMENTO CONCLUÍDO ===")
            self.after(0, lambda: messagebox.showinfo("Sucesso", "Processamento concluído com sucesso!"))
            
        except Exception as e:
            error_msg = f"Erro durante o processamento: {str(e)}"
            self.log_message(f"✗ {error_msg}")
            self.after(0, lambda: messagebox.showerror("Erro", error_msg))
        
        finally:
            # Restaurar interface
            self.after(0, self.finish_processing)
    
    def finish_processing(self):
        """Finaliza o processamento e restaura a interface"""
        self.processing = False
        self.process_button.configure(text="Processar PDFs", state="normal")
        self.progress_bar.set(1.0)
    
    def log_message(self, message):
        """Adiciona mensagem ao log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}"
        
        def update_log():
            self.log_textbox.configure(state="normal")
            self.log_textbox.insert("end", formatted_message + "\n")
            self.log_textbox.configure(state="disabled")
            self.log_textbox.see("end")
        
        if threading.current_thread() == threading.main_thread():
            update_log()
        else:
            self.after(0, update_log)
    
    def on_closing(self):
        """Manipula o fechamento da janela"""
        if self.processing:
            if messagebox.askokcancel("Fechar", "Processamento em andamento. Deseja realmente fechar?"):
                self.destroy()
        else:
            self.destroy()

def main():
    """Função principal"""
    try:
        app = App()
        app.mainloop()
    except Exception as e:
        print(f"Erro ao iniciar a aplicação: {e}")
        import traceback
        traceback.print_exc()
        input("Pressione Enter para sair...")

if __name__ == "__main__":
    main()
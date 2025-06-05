import customtkinter as ctk
from tkinter import filedialog
import os
import glob
import threading
from main import process_pdf_file # Vamos criar essa função em main.py

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- Configuração da Janela Principal ---
        self.title("Extrator de Notas Fiscais")
        self.geometry("700x550")
        self.resizable(False, False)
        ctk.set_appearance_mode("System")  # Pode ser "Dark" ou "Light"
        ctk.set_default_color_theme("blue")

        # --- Variáveis de Estado ---
        self.input_folder_path = ctk.StringVar(value="Nenhuma pasta selecionada")
        self.output_folder_path = ctk.StringVar(value=os.path.join(os.getcwd(), "output"))

        # --- Layout com Grid ---
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1) # Linha do log terá mais espaço

        # --- Frame de Entrada ---
        input_frame = ctk.CTkFrame(self)
        input_frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        input_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(input_frame, text="Pasta de Entrada:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, pady=10)
        ctk.CTkEntry(input_frame, textvariable=self.input_folder_path, state="readonly").grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(input_frame, text="Selecionar Pasta...", command=self.select_input_folder).grid(row=0, column=2, padx=10, pady=10)

        # --- Frame de Filtros e Opções ---
        options_frame = ctk.CTkFrame(self)
        options_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        
        ctk.CTkLabel(options_frame, text="Filtros (Opcional):", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, columnspan=5, padx=10, pady=(10,0), sticky="w")
        ctk.CTkLabel(options_frame, text="Deixe os campos em branco para processar todos os períodos.", font=ctk.CTkFont(size=10, slant="italic")).grid(row=1, column=0, columnspan=5, padx=10, pady=(0,10), sticky="w")

        ctk.CTkLabel(options_frame, text="Mês Inicial:").grid(row=2, column=0, padx=(20, 5), pady=10, sticky="e")
        self.start_month_var = ctk.StringVar(value="1")
        ctk.CTkOptionMenu(options_frame, variable=self.start_month_var, values=[str(i) for i in range(1, 13)]).grid(row=2, column=1, padx=5, pady=10, sticky="w")

        ctk.CTkLabel(options_frame, text="Mês Final:").grid(row=2, column=2, padx=(20, 5), pady=10, sticky="e")
        self.end_month_var = ctk.StringVar(value="12")
        ctk.CTkOptionMenu(options_frame, variable=self.end_month_var, values=[str(i) for i in range(1, 13)]).grid(row=2, column=3, padx=5, pady=10, sticky="w")

        ctk.CTkLabel(options_frame, text="Ano:").grid(row=3, column=0, padx=(20, 5), pady=10, sticky="e")
        self.year_var = ctk.StringVar()
        ctk.CTkEntry(options_frame, textvariable=self.year_var, placeholder_text="Ex: 2024", width=100).grid(row=3, column=1, padx=5, pady=10, sticky="w")

        self.exclude_canceled_var = ctk.BooleanVar()
        ctk.CTkCheckBox(options_frame, text="Excluir notas canceladas", variable=self.exclude_canceled_var).grid(row=3, column=2, columnspan=2, padx=20, pady=10, sticky="w")

        # --- Frame de Saída ---
        output_frame = ctk.CTkFrame(self)
        output_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        output_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(output_frame, text="Salvar em:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, pady=10)
        ctk.CTkEntry(output_frame, textvariable=self.output_folder_path, state="readonly").grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(output_frame, text="Salvar em...", command=self.select_output_folder).grid(row=0, column=2, padx=10, pady=10)

        # --- Área de Log/Status ---
        log_frame = ctk.CTkFrame(self)
        log_frame.grid(row=3, column=0, padx=20, pady=10, sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(0, weight=1)

        self.log_textbox = ctk.CTkTextbox(log_frame, state="disabled", wrap="word")
        self.log_textbox.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # --- Botão de Ação ---
        self.process_button = ctk.CTkButton(self, text="Iniciar Processamento", command=self.start_processing, state="disabled")
        self.process_button.grid(row=4, column=0, padx=20, pady=20, sticky="ew")

    def select_input_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.input_folder_path.set(folder_selected)
            self.log_message(f"Pasta de entrada selecionada: {folder_selected}")
            self.process_button.configure(state="normal")

    def select_output_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.output_folder_path.set(folder_selected)
            self.log_message(f"Pasta de saída selecionada: {folder_selected}")

    def start_processing(self):
        # Inicia o processamento em uma nova thread para não congelar a GUI
        thread = threading.Thread(target=self.processing_thread)
        thread.start()

    def processing_thread(self):
        """
        Thread que executa o processamento dos arquivos para não travar a interface.
        """
        self.process_button.configure(state="disabled")
        self.log_message("Iniciando o processamento...")

        input_path = self.input_folder_path.get()
        output_path = self.output_folder_path.get()

        if not os.path.isdir(input_path):
            self.log_message(f"ERRO: A pasta de entrada não é válida: {input_path}")
            self.process_button.configure(state="normal")
            return

        pdf_files = glob.glob(os.path.join(input_path, "*.pdf"))
        if not pdf_files:
            self.log_message("Nenhum arquivo PDF encontrado na pasta de entrada.")
            self.process_button.configure(state="normal")
            return

        self.log_message(f"Encontrados {len(pdf_files)} arquivos PDF.")

        start_month = self.start_month_var.get()
        end_month = self.end_month_var.get()
        year = self.year_var.get()
        
        # Converte para os tipos corretos antes de passar para a função
        start_month_arg = int(start_month)
        end_month_arg = int(end_month)
        year_arg = int(year) if year.isdigit() else None
        exclude_canceled = self.exclude_canceled_var.get()

        # Validação simples
        if start_month_arg > end_month_arg:
            self.log_message("ERRO: O mês inicial não pode ser maior que o mês final.")
            self.process_button.configure(state="normal")
            return

        for pdf_file in pdf_files:
            try:
                self.log_message(f"Processando: {os.path.basename(pdf_file)}...")
                # Chamada à função de processamento que virá do main.py
                result_path = process_pdf_file(
                    pdf_file,
                    output_path,
                    start_month_arg,
                    end_month_arg,
                    year_arg,
                    exclude_canceled
                )
                if result_path:
                    self.log_message(f"SUCESSO: Arquivo salvo em {result_path}")
                else:
                    self.log_message(f"FALHA: Não foi possível processar {os.path.basename(pdf_file)}")
            except Exception as e:
                self.log_message(f"ERRO ao processar {os.path.basename(pdf_file)}: {e}")

        self.log_message("\nProcessamento concluído!")
        self.process_button.configure(state="normal")

    def log_message(self, message):
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.configure(state="disabled")
        self.log_textbox.see("end")

if __name__ == "__main__":
    app = App()
    app.mainloop()
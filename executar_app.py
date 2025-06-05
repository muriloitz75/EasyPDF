"""
Script principal para iniciar a interface gráfica da aplicação.
"""
import sys
import subprocess

try:
    # Tenta importar o customtkinter para verificar se está instalado
    import customtkinter
except ImportError:
    print("A biblioteca 'customtkinter' não está instalada.")
    print("Instalando dependências a partir de requirements.txt...")
    
    # Tenta instalar as dependências usando pip
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("\nDependências instaladas com sucesso!")
        print("Por favor, execute a aplicação novamente.")
    except subprocess.CalledProcessError:
        print("\nERRO: Falha ao instalar as dependências.")
        print("Por favor, instale manualmente executando: pip install -r requirements.txt")
    
    # Sai para o usuário poder executar novamente
    sys.exit(1)

# Se a importação foi bem-sucedida, inicia a aplicação
from gui import App

if __name__ == "__main__":
    try:
        app = App()
        app.mainloop()
    except Exception as e:
        print(f"Ocorreu um erro inesperado ao executar a aplicação: {e}")
        input("Pressione ENTER para sair.")

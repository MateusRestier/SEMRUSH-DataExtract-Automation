import subprocess
import os

def executar_script(script_name):
    """Executa um script Python localizado no mesmo diretório."""
    try:
        script_path = os.path.join(os.path.dirname(__file__), script_name)
        result = subprocess.run(['python', script_path], check=True, text=True)
        print(f"{script_name} executado com sucesso.")
    except subprocess.CalledProcessError as e:
        print(f"Erro ao executar {script_name}: {e}")
    except Exception as e:
        print(f"Erro inesperado ao executar {script_name}: {e}")

def main():
    scripts = [
        'login.py',
        'automacaoSEMRUSH.py',
        'tratamentotabelas.py',
        'jogar pro banco.py'
    ]

    for script in scripts:
        executar_script(script)

if __name__ == "__main__":
    main()

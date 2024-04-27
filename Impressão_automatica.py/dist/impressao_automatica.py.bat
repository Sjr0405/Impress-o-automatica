import os
import subprocess
import time
import send2trash
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Pasta que será monitorada
folder_to_watch = r"\\SERVER_PC\Users\Public\Pictures\imprimir"

# Caminho para o Apache OpenOffice soffice.exe
apache_openoffice_path = r"C:\Program Files\OpenOffice 4\program\soffice.exe"

# Diretório de saída para os arquivos convertidos
output_directory = r"C:\Users\Server\Pictures\output_log"

# Função para imprimir o documento usando o Apache OpenOffice
def imprimir_documento_apache_openoffice(file_path):
    try:
        # Comando para imprimir usando o Apache OpenOffice
        comando = [
            apache_openoffice_path, 
            "--headless", 
            "--convert-to", "pdf", 
            "--outdir", output_directory, 
            file_path
        ]
        subprocess.run(comando, shell=True)
        print("Arquivo impresso com sucesso!")

        # Move o arquivo para a lixeira após a impressão
        send2trash.send2trash(file_path)
        print("Arquivo movido para a lixeira após impressão.")
    except Exception as e:
        print("Erro ao imprimir arquivo:", e)

# Classe para manipular os eventos de criação e movimentação de arquivo
class MyHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory:  # Verifica se o evento não é para uma nova pasta
            file_path = event.src_path
            print("Novo arquivo detectado:", file_path)
            imprimir_documento_apache_openoffice(file_path)

    def on_modified(self, event):
        if not event.is_directory:  # Verifica se o evento não é para uma nova pasta
            file_path = event.src_path
            print("Arquivo modificado detectado:", file_path)
            imprimir_documento_apache_openoffice(file_path)

    def on_moved(self, event):
        if not event.is_directory:  # Verifica se o evento não é para uma nova pasta
            file_path = event.dest_path
            print("Arquivo movido detectado:", file_path)
            imprimir_documento_apache_openoffice(file_path)

# Função principal para rodar como script
def main_script():
    event_handler = MyHandler()
    observer = Observer()
    observer.schedule(event_handler, folder_to_watch, recursive=True)
    observer.start()
    try:
        print("Aguardando novos documentos...")
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

if __name__ == "__main__":
    main_script()  # Executar como script

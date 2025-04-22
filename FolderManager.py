import os
from pathlib import Path
import shutil
import zipfile
import subprocess

class FolderManager:

    def extract_from_folder(self, pasta):
        for item in pasta.iterdir():
            if item.is_dir():
                try:
                    print(f"[Pasta] {item.name}")

                    for subitem in item.iterdir():
                        try:
                            destino = pasta / subitem.name
                            shutil.move(str(subitem), destino)

                        ## caso o nome do item que está na pasta já exista no destino 
                        except:
                            # shutil.move(str(subitem), destino)
                            pass

                    # ## Apaga de forma forçada
                    # shutil.rmtree(item)

                    # Só funciona se a pasta estiver vazia
                    item.rmdir()
                
                except:
                    pasta
    
    def extract_from_zip(self, pasta):
        """
            Percorre uma pasta e extrai todos os arquivos ZIP , excluindo-os após a extração.
        """
        try:
            for file_name in os.listdir(pasta):
                file_path = os.path.join(pasta, file_name)

                extract_to = os.path.splitext(file_path)[0]
            
                if file_name.lower().endswith('.zip'):
                    print(f"[.zip] {file_name}")

                    os.makedirs(extract_to, exist_ok=True)

                    with zipfile.ZipFile(file_path, 'r') as archive:
                        archive.extractall(extract_to)
                        
                    os.remove(file_path)
        
        except Exception as e:
            print("Error > FolderManager > extract_from_zip: ", e)
                        
    def extract_from_rar(self, pasta):
        """
            Percorre uma pasta e extrai todos os arquivos RAR , excluindo-os após a extração.
        """
       
        for file_name in os.listdir(pasta):
            file_path = os.path.join(pasta, file_name)

            extract_to = os.path.splitext(file_path)[0]

            if file_name.lower().endswith('.rar'):
                try:
                    print(f"[.rar] {file_name}")

                    # cria a pasta caso ela não exista
                    os.makedirs(extract_to, exist_ok=True)

                    subprocess.run(["UnRAR.exe", "x", "-o+", file_path, extract_to], 
                                    check=True,
                                    stdout=subprocess.DEVNULL,
                                    stderr=subprocess.DEVNULL
                                    )
                    
                    # remove a pasta rar
                    os.remove(file_path)

                except FileNotFoundError:
                    print("Erro: 'unrar' não encontrado. Certifique-se de que está instalado e no PATH.")
                
                except subprocess.CalledProcessError:
                    print(f"Erro ao extrair {file_path}.")

    def separates_excel_files(self, source_folder, destination_folder):
        """
            Move arquivos Excel de uma pasta para outra.
            
            :param source_folder: Pasta de origem.
            :param destination_folder: Pasta de destino.
        """
        # Verifica se a pasta de destino existe, senão cria
        if not os.path.exists(destination_folder):
            os.makedirs(destination_folder)
            
        # Lista todos os arquivos Excel na pasta
        excel_files = [f for f in os.listdir(source_folder) if f.lower().endswith(('.xlsx', '.xls', 'xlsb', 'xlsm'))]
        
        if not excel_files:
            print("Nenhum arquivo Excel encontrado na pasta.")
            return
        
        for excel_file in excel_files:
            source_path = os.path.join(source_folder, excel_file)
            destination_path = os.path.join(destination_folder, excel_file)
            
            try:
                shutil.move(source_path, destination_path)
                print(f"Arquivo {excel_file} movido para {destination_folder}")

            except Exception as e:
                print(f"Erro ao mover {excel_file}: {e}")
from pathlib import Path
from FolderManager import FolderManager

# caminho = 'C:/Users/patricia.berti.sutto/Downloads/Lote 242'  # Ex: '/home/usuario/documentos'
# caminho = 'C:/Users/patricia.berti.sutto/Downloads/teste'
caminho = 'C:/Users/patricia.berti.sutto/Downloads/Lote 24'
pasta = Path(caminho)  # agora é um objeto Path

fm = FolderManager()


# fm.extract_from_folder(pasta)
# fm.extract_from_zip(pasta)
counter = 0 

while (counter < 3):
    print('r ', counter)
    counter += 1
    fm.extract_from_folder(pasta)
    fm.extract_from_zip(pasta)

    fm.extract_from_rar(pasta)

    fm.extract_from_folder(pasta)


## separa os arquivos excel
source_folder = 'C:/Users/patricia.berti.sutto/Downloads/Lote 24'
destination_folder = 'C:/Users/patricia.berti.sutto/Downloads/just_excel_L24'

fm.separates_excel_files(source_folder, destination_folder)
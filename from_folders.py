import os
from pathlib import Path
import shutil

caminho = 'C:/Users/patricia.berti.sutto/Downloads/Lote 222'  # Ex: '/home/usuario/documentos'
pasta = Path(caminho)  # agora é um objeto Path

for item in pasta.iterdir():
    if item.is_dir():
        try:
            print(f"[Pasta] {item.name}")

            for subitem in item.iterdir():
            
                destino = pasta / subitem.name
                shutil.move(str(subitem), destino)

            shutil.rmtree(item)  # Só funciona se a pasta estiver vazia
        
        except:
            pasta
    
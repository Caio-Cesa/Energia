import os
import shutil

base_path = r"c:\Users\Usuario\OneDrive\Documentos\Valerio"
dest_dir = os.path.join(base_path, "Dados_Entrada")

if not os.path.exists(dest_dir):
    os.makedirs(dest_dir)
    print(f"Created directory: {dest_dir}")

for i in range(1, 8):
    folder_name = f"Ralatorio_Emerg{i}"
    src_dir = os.path.join(base_path, folder_name)
    
    if os.path.exists(src_dir):
        print(f"Processing folder: {folder_name}")
        for filename in os.listdir(src_dir):
            if filename.endswith(".txt"):
                # Construct new filename: "OldName i.txt"
                name_part, ext_part = os.path.splitext(filename)
                new_filename = f"{name_part} {i}{ext_part}"
                
                src_file = os.path.join(src_dir, filename)
                dest_file = os.path.join(dest_dir, new_filename)
                
                shutil.move(src_file, dest_file)
                print(f"  Moved and renamed: {filename} -> {new_filename}")
    else:
        print(f"Folder not found: {folder_name}")

print("Data reorganization complete.")

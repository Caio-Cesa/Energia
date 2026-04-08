import os

file_path = r"c:\Users\Usuario\OneDrive\Documentos\Valerio\VBA.bas"
with open(file_path, "rb") as f:
    content = f.read()

# Try to decode
try:
    decoded = content.decode('utf-16')
except UnicodeDecodeError:
    try:
        decoded = content.decode('utf-8')
    except UnicodeDecodeError:
        decoded = content.decode('windows-1252', errors='ignore')

normalized = decoded.replace('\r\n', '\n').replace('\r', '\n')
print(normalized)

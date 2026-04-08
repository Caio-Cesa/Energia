import os

file_path = r"c:\Users\Usuario\OneDrive\Documentos\Valerio\VBA.bas"
with open(file_path, "rb") as f:
    content = f.read()

print(f"File size: {len(content)} bytes")
print(f"First 100 bytes (hex): {content[:100].hex()}")

# Try to decode and normalize line endings
try:
    decoded = content.decode('utf-16')
    print("Detected UTF-16 encoding")
except UnicodeDecodeError:
    try:
        decoded = content.decode('utf-8')
        print("Detected UTF-8 encoding")
    except UnicodeDecodeError:
        decoded = content.decode('windows-1252', errors='ignore')
        print("Falling back to windows-1252")

normalized = decoded.replace('\r\n', '\n').replace('\r', '\n')
lines = normalized.split('\n')
for i, line in enumerate(lines[:150]): # Read more lines
    print(f"{i+1:3}: {line}")

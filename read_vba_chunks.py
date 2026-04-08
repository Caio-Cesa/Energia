import os

file_path = r"c:\Users\Usuario\OneDrive\Documentos\Valerio\VBA.bas"
with open(file_path, "rb") as f:
    content = f.read()

try:
    decoded = content.decode('utf-16')
except UnicodeDecodeError:
    try:
        decoded = content.decode('utf-8')
    except UnicodeDecodeError:
        decoded = content.decode('windows-1252', errors='ignore')

normalized = decoded.replace('\r\n', '\n').replace('\r', '\n')
lines = normalized.split('\n')

def print_range(start, end):
    print(f"--- Lines {start} to {end} ---")
    for i in range(start-1, min(end, len(lines))):
        print(f"{i+1:4}: {lines[i]}")

print_range(1, 10)
print_range(11, 60)
print_range(61, 112)
print_range(113, 160)
print_range(161, 210)

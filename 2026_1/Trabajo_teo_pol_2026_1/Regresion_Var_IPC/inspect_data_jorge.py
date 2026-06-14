import openpyxl

path = r'C:\Users\51950\Documents\ULIMA\2026_1\Econometria 2\Trabajo_metria_2_2026_1\Data\DATA_JORGE.xlsx'
wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
ws = wb['Sheet1']

rows = list(ws.iter_rows(values_only=True))
print("Fila 1 (headers):", rows[0])
print("Fila 2 (unidades):", rows[1])
print("Fila 3 (1er dato):", rows[2])
print("Fila 4:", rows[3])
print("...")
print(f"Total filas: {len(rows)}")
print("Ultima fila:", rows[-1])
print("Penult fila:", rows[-2])
print("Antepenult:", rows[-3])

# Contar no-nulos por columna (indices 0 a 13)
headers = rows[0]
for ci, h in enumerate(headers):
    vals = [r[ci] for r in rows[2:] if r[ci] is not None]
    print(f"  Col {ci} [{h}]: {len(vals)} no-nulos, primer={vals[0] if vals else 'N/A'}, ultimo={vals[-1] if vals else 'N/A'}")

wb.close()

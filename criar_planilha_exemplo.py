import pandas as pd

entrada = pd.DataFrame({
    'material': [
        'Bomb. pneum s/bc pt',
        'tubo galv 1/2',
        'paraf sext 1/4 x 20',
    ],
    'cadastro sugerido': ['', '', ''],
    'codigo': ['', '', ''],
    'ncm sugerido': ['', '', ''],
    'tipo': ['', '', ''],
    'status': ['', '', ''],
    'resultado': ['', '', ''],
})

banco = pd.DataFrame({
    'codigo': ['001234', '005678'],
    'material': [
        'BOMBA PNEUMATICA SEM BOCA PRETA',
        'TUBO GALVANIZADO 1 POLEGADA',
    ],
    'ncm': ['84137010', '73063000'],
    'tipo': ['UN', 'MT'],
})

with pd.ExcelWriter('planilha.xlsx', engine='openpyxl') as writer:
    entrada.to_excel(writer, sheet_name='Sheet1', index=False)
    banco.to_excel(writer, sheet_name='banco', index=False)

print('planilha.xlsx criada.')

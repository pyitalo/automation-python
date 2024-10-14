import openpyxl as op

arquivo = 'planilha_fechamento.xlsx'
planilha_fechamento = op.load_workbook(arquivo)
sheet = planilha_fechamento.active
# seleciona a planilha
for linha in sheet.iter_rows():
    for bloco in linha:
        bloco.value = None
planilha_fechamento.save(arquivo)
print('Arquivo limpo com sucesso!')

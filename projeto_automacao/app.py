import openpyxl
from PIL import Image,ImageDraw, ImageFont

workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for indice ,linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    nome_curso = linha[0].value
    nome_aluno = linha[1].value
    tipo_participacao = linha[2].value
    data_inicio = linha[3].value
    data_termino = linha[4].value
    carga_horaria = linha[5].value
    data_emissao_certificado = linha[6].value

    Font_image = ImageFont.truetype('./tahomabd.ttf',90)
    Font_geral = ImageFont.truetype('./tahoma.ttf',80)
    font_data = ImageFont.truetype('./tahoma.ttf',55)
    
    image = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)
    
    desenhar.text((1020,815), nome_aluno, fill='black', font=Font_image)
    desenhar.text((1100,950), nome_curso, fill='black', font=Font_geral)
    desenhar.text((1460,1058), tipo_participacao, fill='black', font=Font_geral)
    desenhar.text((1485, 1182), str(carga_horaria), fill='black', font=Font_geral)
    
    desenhar.text((750, 1770), data_inicio,fill='black', font=font_data)
    desenhar.text((750, 1930), data_termino,fill='black', font=font_data)
    desenhar.text((2220, 1930), data_emissao_certificado,fill='black', font=font_data)
    
    image.save(f'./{indice}-{nome_aluno} certificado.png')
    


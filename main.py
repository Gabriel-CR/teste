from reportlab.pdfgen import canvas
import pandas as pd
from PIL import Image
import sys

# Converte mm em points (unidade de medida usado na biblioteca reportlab)
def mmToPoint(mm):
    return mm / 0.352777

# Retorna uma tupla com lista de alunos e nome do diretor, obtidos de um arquivo excell
def getAlunosEDiretor(caminhoPlanilha):
    lines = pd.read_excel(caminhoPlanilha)
    listaAlunos = []

    for index, aluno in lines.iterrows():
        listaAlunos.append(aluno['nome'])

    diretor = lines.iloc[0]['diretor']

    return (listaAlunos, diretor)

# Obtem dimensões da imagem do certificado
def getTamanhoImagem(caminhoImg):
    img = Image.open(caminhoImg) 
    largura = img.width 
    altura = img.height
    
    return (largura, altura)

'''
    Obtem caminho da imagem padrão do certificado
    Obtem largura e altura do certificado
    Obtem alunos para preencher o certificado
    Obtem nome do diretor
    Retorna um dicionário com todos os dados para a emissão dos certificados
'''
def lerDadosPdf(certificado, planilha):
    (largura, altura) = getTamanhoImagem(certificado)
    (alunos, diretor) = getAlunosEDiretor(planilha)

    dicionario = {
        'size': (largura, altura),
        'alunos': alunos,
        'diretor': diretor,
        'img': certificado
    }

    return dicionario

'''
    Criar os certificados com fonte padrão
    Cria um certificado para cada aluno
    Todos os certificados são acompanados do nome do diretor
    As coordenadas para preenchimento do certifidado só são válidas para a imagem que acompanha o código
'''
def gerarCertificados(alunos, size, caminhoImgCertificado, diretor):
    for aluno in alunos:
        pdf = canvas.Canvas('./' + aluno + '.pdf', pagesize=size)
        pdf.drawImage(caminhoImgCertificado, 0, 0)

        pdf.setFont('Helvetica-Oblique', 80)
        pdf.drawCentredString(mmToPoint(355), mmToPoint(300), aluno)

        pdf.setFont('Helvetica', 24)
        pdf.drawCentredString(mmToPoint(168), mmToPoint(115), aluno)
        pdf.drawCentredString(mmToPoint(538), mmToPoint(115), diretor)
        
        pdf.save()

if __name__ == "__main__":
    caminhoImgCertificado, caminhoPlanilha = sys.argv[1], sys.argv[2]
    dados = lerDadosPdf(caminhoImgCertificado, caminhoPlanilha)
    gerarCertificados(dados['alunos'], dados['size'], dados['img'], dados['diretor'])
from pdf2docx import Converter
import os

def converter_pdf_para_docx(pasta_pdf, pasta_saida):
    
    pasta_pdf = os.path.abspath(pasta_pdf)
    pasta_saida = os.path.abspath(pasta_saida)
    

    if not os.path.exists(pasta_saida):
        os.makedirs(pasta_saida)

    for arquivo in os.listdir(pasta_pdf):
        if arquivo.lower().endswith('.pdf'):
            caminho_completo_pdf = os.path.join(pasta_pdf, arquivo)
            caminho_completo_docx = os.path.join(pasta_saida, f"{os.path.splitext(arquivo)[0]}.docx")
            

            cv = Converter(caminho_completo_pdf)
         
            cv.convert(caminho_completo_docx)
            
           
            cv.close()
            
            print(f"Convertido: {arquivo} para DOCX")


pasta_pdf = '/Users/rafaelives/Desktop/arv/bradesco'
pasta_saida = 'Users/rafaelives/Desktop/arv'

converter_pdf_para_docx(pasta_pdf, pasta_saida)

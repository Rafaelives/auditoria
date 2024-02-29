from pdf2docx import Converter
import os

def converter_pdf_para_docx(pasta_pdf, pasta_saida):
    # Garante que os caminhos das pastas são absolutos
    pasta_pdf = os.path.abspath(pasta_pdf)
    pasta_saida = os.path.abspath(pasta_saida)
    
    # Certifica-se de que a pasta de saída existe
    if not os.path.exists(pasta_saida):
        os.makedirs(pasta_saida)
    
    # Lista todos os arquivos PDF na pasta
    for arquivo in os.listdir(pasta_pdf):
        if arquivo.lower().endswith('.pdf'):
            caminho_completo_pdf = os.path.join(pasta_pdf, arquivo)
            caminho_completo_docx = os.path.join(pasta_saida, f"{os.path.splitext(arquivo)[0]}.docx")
            
            # Cria uma instância do conversor
            cv = Converter(caminho_completo_pdf)
            
            # Converte o PDF para DOCX
            cv.convert(caminho_completo_docx)
            
            # Fecha o conversor
            cv.close()
            
            print(f"Convertido: {arquivo} para DOCX")

# Substitua os caminhos abaixo conforme necessário
pasta_pdf = '/Users/rafaelives/Desktop/escalas'
pasta_saida = 'Users/rafaelives/Desktop/arv'

converter_pdf_para_docx(pasta_pdf, pasta_saida)

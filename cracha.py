import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from tabulate import tabulate
import requests
from PIL import Image, ImageFile
from bs4 import BeautifulSoup
import PIL
from fpdf import FPDF
import re
from dotenv import load_dotenv

# Carregar variáveis de ambiente do arquivo .env
load_dotenv()
# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = os.getenv("SAMPLE_SPREADSHEET_ID")
SAMPLE_RANGE_NAME = os.getenv("SAMPLE_RANGE_NAME")

LARGURA = 100
ALTURA = 145

def extract_dates(dias):
    # Expressão regular para encontrar datas no formato "dd de Mês"
    pattern = r'\d{1,2}'
    return re.findall(pattern, dias)

def main():
  """Shows basic usage of the Sheets API.
  Prints values from a sample spreadsheet.
  """
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "client_secret.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())


  try:
    service = build("sheets", "v4", credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = (
        sheet.values()
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME)
        .execute()
    )
    values = result.get("values", [])

    if not values:
      print("No data found.")
      return

    # Crie um novo objeto PDF
    pdf = FPDF("P","mm",(LARGURA,ALTURA))
    
    
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36'}

    # Itere pelas linhas da planilha
    for row in values:
        cargo1 = row[2]
        cargo2 = row[3]
        grupo = row[4]
        nome = row[5]
        sobrenome = row[6]
        cidade = row[7]
        rg = row[8]
        imagem_url = row[10]
        dias = row[11] if len(row) > 11 else ""
        id_imagem = imagem_url.split("id=")[1]
        imagem_url = f"https://drive.google.com/uc?export=view&id={id_imagem}"
        imagemtemp = nome+sobrenome

        print(f"Gerando: {nome, sobrenome}\n")
        
        # Extraia apenas as datas
        datas = extract_dates(dias)
        datas_formatadas = ", ".join(datas)
        print(datas_formatadas)
        
        # Baixe a imagem do Google Drive
        imagem_response = requests.get(imagem_url, headers=headers)
        imagem_response.raise_for_status()

        # Verifique o tipo de conteúdo da imagem
        content_type = imagem_response.headers.get('content-type')
        print(content_type)

        # Salve a imagem em um arquivo temporário
        with open(f"{imagemtemp}.jpg", "wb") as f:
            f.write(imagem_response.content)
            
        # Abra a imagem usando o arquivo temporário
        try:
            imagem = Image.open(f"{imagemtemp}.jpg")
        except PIL.UnidentifiedImageError:
            print(f"Erro ao abrir a imagem: {imagem_url}")
            # Tente abrir a imagem com uma extensão diferente
            try:
                imagem = Image.open(f"{imagemtemp}.png")
            except PIL.UnidentifiedImageError:
                print(f"Erro ao abrir a imagem (formato inválido): {imagem_url}")
                continue  # Pula para a próxima imagem se esta não abrir
        except Exception as e:
            print(f"Erro ao abrir a imagem: {e}")
            continue  # Pula para a próxima imagem se esta não abrir

        # Ajuste o tamanho da imagem para 3x4cm
        imagem = imagem.resize((114, 152))
        
        # Adicione os dados ao PDF
        pdf.add_page()
        pdf.image(f"modelocracha.png", x=0,y=0, w=LARGURA, h=ALTURA)
        
        
        # posicão dos textos do crachá
        x_text = 7
        y_text = 100
        kerningNome = 5
        kerning = 4.5

        # Categoria
        # Nome Grupo

        # Adiciona o texto e seta o kerning
        # (NOME)
        pdf.set_font("Arial", 'B', 12)
        pdf.text(x_text, y_text, f"{nome.upper()}")
        y_text = y_text + kerningNome
        pdf.text(x_text, y_text, f"{sobrenome.upper()}")
        y_text = y_text + (kerning*2)
        

        # Restante dos textos
        pdf.set_font("Arial","",9)
        pdf.text(x_text, y_text, f"{cidade.upper()}")
        y_text = y_text + kerning
        pdf.text(x_text, y_text, f"RG: {rg.upper()}")
        y_text = y_text + (kerning*2)

        pdf.text(x_text, y_text, cargo1.upper())
        if cargo2 != "":
           y_text = y_text + kerning
           pdf.text(x_text, y_text, f"e {cargo2.upper()}")
        
        pdf.set_font("Arial","B",9)
        y_text = y_text + kerning
        pdf.text(x_text, y_text, grupo.upper())
        y_text = y_text + kerning
        if datas_formatadas != "":
           pdf.text(x_text, y_text, f"Dias: {datas_formatadas}/JUL")
                    
        

        # Adicione a imagem ao PDF
        imagem.save(f"{imagemtemp}.jpg")
        pdf.image(f"{imagemtemp}.jpg", x=57,y=91, w=35, h=46.7)
       

        # Remova o arquivo de imagem temporário
        os.remove(f"{imagemtemp}.jpg")

    # Salve o arquivo PDF
    pdf.output("output.pdf")

    print("Arquivo PDF criado com sucesso!")

  except HttpError as err:
    print(err)


if __name__ == "__main__":
  main()
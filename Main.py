import pdfplumber
import re
import math
from docxtpl import DocxTemplate
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches, Cm

rotulos = ['CEP','Nome','Endereco','Telefone','Email','Qtd_mod','Qtd_inv','Pot_nom']
rotulos_pdf = ['CEP da UC com GD','Nome do Titular da UC com GD','Endereço','Telefone do Titular \(DDD \+ número\)','E-mail do Titular da UC com GD','Quantidade de Módulos','Quantidade de Inversores','Potência Total dos Módulos \(kW\)']
energia_mensal = {'Energia_jan':0,'Energia_fev':0,'Energia_mar':0,'Energia_abr':0,'Energia_maio':0,'Energia_jun':0,'Energia_jul':0,'Enegia_ago':0,'Energia_set':0,'Energia_out':0,'Energia_nov':0,'Energia_dez':0}
irradiacao = [5.01, 5.5, 5.1, 5.46, 5.56, 5.61, 5.83, 6.47, 5.91, 5.45, 4.75, 4.98]
mes = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

gerador_trina = {'Fabricante':'Trina Solar',
                 'SIGLA':'TSM-695NEG21C.20',
                 'Tec_construcao':'Monocristalino',
                 'Garantia':'12 anos',
                 'Pot_max':'695 W',
                 'Eficiencia':'22,4 %',
                 'Tensao_nom':'40,3 V',
                 'Tensao_aberto':'48,3 V',
                 'Corrente_nom':'17,25 A',
                 'Corrente_cc':'18,28 A',
                 'axlxp':'2384 x 1303 x 33 mm',
                 'Peso':'38,3 kg',
                 'Imagem_gerador':'Imagens/Trina_gerador.jpg'}

gerador_astronergy = {
    'Fabricante': 'Astronergy',
    'SIGLA': 'CHSM6612M/HV - 375W',
    'Tec_construcao': 'Monocristalino',
    'Garantia': '10 anos',
    'Pot_max': '375 W',
    'Eficiencia': '19,4 %',
    'Tensao_nom': '39,76 V',
    'Tensao_aberto': '48,45 V',
    'Corrente_nom': '9,45 A',
    'Corrente_cc': '9,94 A',
    'axlxp': '1960 x 992 x 40 mm',
    'Peso': '21,8 kg',
    'Imagem_gerador':'Imagens/Trina_gerador.jpg'
}

gerador_canadian = {'Fabricante':'Canadian Solar',
                    'SIGLA':'TCS6U-330P',
                    'Tec_construcao':'Policristalino',
                    'Garantia':'12 anos',
                    'Pot_max':'330 W',
                    'Eficiencia':'16,97 %',
                    'Tensao_nom':'37,2 V',
                    'Tensao_aberto':'45,6 V',
                    'Corrente_nom':'8,88 A',
                    'Corrente_cc':'9,45 A',
                    'axlxp':'1990 x 992 x 40 mm',
                    'Peso':'22,4 kg',
                    'Imagem_gerador':'Imagens/Trina_gerador.jpg'
                    }

gerador_gokin = {
    'Fabricante': 'Gokin',
    'SIGLA': 'GK-1-72HT585M',
    'Tec_construcao': 'Monocristalino',
    'Garantia': '12 anos',
    'Pot_max': '585 W',
    'Eficiencia': '22,6%',
    'Tensao_nom': '42.74 V',
    'Tensao_aberto': '51.67 V',
    'Corrente_nom': '13.69 A',
    'Corrente_cc': '14.43 A',
    'axlxp': '2310x1125x1259mm',
    'Peso': '26,8 kg',
    'Imagem_gerador':'Imagens/Trina_gerador.jpg'
}

gerador_hanersun_585 = {
    'Fabricante': 'Hanersun',
    'SIGLA': 'HN18-72H585',
    'Tec_construcao': 'Monocristalino',
    'Garantia': '12 anos',
    'Pot_max': '585 W',
    'Eficiencia': '22,65%',
    'Tensao_nom': '40,80 V',
    'Tensao_aberto': '49,30 V',
    'Corrente_nom': '10,91 A',
    'Corrente_cc': '11,53 A',
    'axlxp': '2278*1134*30mm',
    'Peso': '28,5 kg',
    'Imagem_gerador':'Imagens/Trina_gerador.jpg'
}

inversor_growatt2 = {
    'Fabricante_sigla': 'Growatt NEO 2000M-X',
    'Entradas': '4',
    'Monitoramento': 'SIM – Wireless',
    'Pot_nom_max': '2 kW',
    'Tensao_nom_freq': '220 V - 54/65 Hz',
    'Tensao_max': '65 VCC',
    'Tensao_saida': '160 – 285 V',
    'Corrente_max_saida': '9,3 A',
    'Eficiencia_max': '96,5 %',
    'axlxp_inv': '396 × 300 × 45 mm',
    'Peso_inv': '5 kg',
    'Nome_inversor': 'Growatt',
    'Link_inversor': '',
    'App_inversor': '',
    'Imagem_inversor': 'Imagens/inversor_growatt_2kw.jpg',
}

inversor_growatt225 = {'Fabricante_sigla':'Growatt NEO 2250M-X2',
                       'Entradas':'4',
                       'Monitoramento':'SIM – Wireless',
                       'Pot_nom_max':'2.25 kW',
                       'Tensao_nom_freq':'35 V - 50/60 Hz',
                       'Tensao_max':'60 VCC',
                       'Tensao_saida':'220 V',
                       'Corrente_max_saida':'10,23 A',
                       'Eficiencia_max':'96,5%',
                       'axlxp_inv':'396*270*45 mm',
                       'Peso_inv':'5,1 kg',
                       'Nome_inversor':'Growatt',
                       'Link_inversor':'https://server.growatt.com/login',
                       'App_inversor':'ShinePhone',
                       'Imagem_inversor':'Imagens/inversor_growatt.jpg'}

inversor_Sungrow = {'Fabricante_sigla':'Sungrow SG3K-S',
                    'Entradas':'1',
                    'Monitoramento':'SIM – Wireless',
                    'Pot_nom_max':'3 kW',
                    'Tensao_nom_freq':'220 V - 60 Hz',
                    'Tensao_max':'600 VCC',
                    'Tensao_saida':'176 - 276 V',
                    'Corrente_max_saida':'13,7 A',
                    'Eficiencia_max':'98,2 %',
                    'axlxp_inv':'370 x 300 x 125 mm',
                    'Peso_inv':'8,5 kg'}

inversor_Hoymiles = {
    'Fabricante_sigla': 'Hoymiles MI-1500 / MI-700',
    'Entradas': '2',
    'Monitoramento': 'SIM – Wireless',
    'Pot_nom_max': '1,2 kW / 0,7 kW',
    'Tensao_nom_freq': '220 V - 45/65 Hz',
    'Tensao_max': '60 Vcc',
    'Tensao_saida': '180 - 275 V',
    'Corrente_max_saida': '5,21 A / 3,36 A',
    'Eficiencia_max': '96,50 % / 96,70%',
    'axlxp_inv': '176 x 280 x 33 mm',
    'Peso_inv': '3,75 kg',
    'Imagem_inversor':'Imagens/inversor_hoymiles.jpg'
}

inversor_hyxipower_m2000 = {
    'Fabricante_sigla': 'Hyxipower HYX-M2000-SW',
    'Entradas': '4',
    'Monitoramento': 'SIM – Wireless',
    'Pot_nom_max': '2 kW',
    'Tensao_nom_freq': '220 V - 50/60 Hz',
    'Tensao_max': '65 VCC',
    'Tensao_saida': '183 – 276 V',
    'Corrente_max_saida': '9,09 A',
    'Eficiencia_max': '96,70%',
    'axlxp_inv': '310*236*35.5mm',
    'Peso_inv': '5 kg',
    'Nome_inversor': 'Hyxipower',
    'Link_inversor': '',
    'App_inversor': '',
    'Imagem_inversor': 'Imagens/inversor_hyxipower.jpg',
}
dicionario = {}

print("Escolha o tipo de módulo:")
print("[1] Trina Solar (TSM-695NEG21C.20)")
print("[2] Canadian Solar (TCS6U-330P)")
print("[3] Astronergy (CHSM6612M/HV - 375W)")
print("[4] Gokin (GK-1-72HT585M)")
print("[5] Hanersun (HN18-72H585)")
tipo_mod = int(input(''))

print("Escolha o tipo de inversor:")
print("[1] Growatt (NEO 2000M-X ou NEO 2250M-X2)")
print("[2] Sungrow (SG3K-S)")
print("[3] Hoymiles (MI-1500 / MI-700)")
print("[4] Hyxipower (HYX-M2000-SW)")
tipo_inv = int(input(''))

print("Nome: ")
nome = input('')

with pdfplumber.open(f"D:/UnB/Python Stuff/{nome}.pdf") as pdf:
    first_page = pdf.pages[0]
    pdf_text = first_page.extract_text()
    #print (pdf_text)

def valor(label, texto):
    padrao = rf'{label}:\s*(.+)'
    encontrado = re.search(padrao, texto, re.IGNORECASE)
    return encontrado.group(1).strip()

def valor_coordenada(label, texto):
    padrao = rf'{label}\s*(.+)'
    encontrado = re.search(padrao, texto, re.IGNORECASE | re.MULTILINE)
    return encontrado.group(1).strip()

latitude = valor_coordenada('latitude',pdf_text)
longitude = valor_coordenada('longitude',pdf_text)

def gms_para_decimal(coordenada_gms_string):
    padrao= r'(\D)\s*(\d+)\s+(\d+)\s+([\d.,]+)'
    match = re.search(padrao, coordenada_gms_string, re.IGNORECASE)
    
    if not match:
        print(f"Erro: Formato de coordenada inválido para '{coordenada_gms_string}'. Regex falhou.")
        # Retorna N/A se nenhum padrão funcionar.
        return "N/A"

    direcao = match.group(1).upper() # S
    grau = int(match.group(2))       # 15
    minuto = int(match.group(3))     # 51
    segundo = float(match.group(4).replace(',', '.')) # 19.631

    valor_decimal = grau + (minuto / 60) + (segundo / 3600)

    if direcao in ['S', 'O', 'W']:
        valor_decimal *= -1

    return round(valor_decimal, 6)

latitude_decimal = gms_para_decimal(latitude)
longitude_decimal = gms_para_decimal(longitude)

i = 0
for rotulo in rotulos:
    dicionario[rotulo]=valor(rotulos_pdf[i],pdf_text)
    i += 1

Pot_ano = float(dicionario['Pot_nom'].replace(',','.')) * 128 * 12
Pot_ano_rounded = int(Pot_ano // 1)
Pot_ano_rounded = Pot_ano_rounded - (Pot_ano_rounded%100)
Pot_ano_rounded_str = str(Pot_ano_rounded)
dicionario['Pot_ano_rounded'] = Pot_ano_rounded_str

dicionario['Area_mod']= valor('Área Total dos Arranjos \(m²\)',pdf_text)
Qtd_mod_var = int(dicionario['Qtd_mod'])
Area_mod_total = float (dicionario['Area_mod'].replace(',','.')) * Qtd_mod_var
dicionario['Area_mod_total'] = Area_mod_total

dicionario['Latitude'] = latitude_decimal
dicionario['Longitude'] = longitude_decimal

Bairro_partes = dicionario['Endereco'].rsplit(',',1)
Bairro = Bairro_partes[1].strip()
dicionario['Bairro'] = Bairro

Pot_mensal = int((Pot_ano_rounded/12)//1)
dicionario['Pot_mensal'] = Pot_mensal
Pot_diaria = int((Pot_mensal / 30)//1)
dicionario['Pot_diaria'] = Pot_diaria

Nomes = dicionario['Nome'].split()
Nome_capa = f'{Nomes[0]} {Nomes[-1]}'
dicionario ['Nome_capa'] = Nome_capa
Nome_login = f'{Nomes[0]}{Nomes[-1]}'
dicionario['Nome_login'] = Nome_login
Senha_login = f'{Nomes[0][0]}{Nomes[-1][0]}123456'
Senha_login = Senha_login.lower()
dicionario['Senha_login'] = Senha_login

#Seleção dos módulos e  inversores
if tipo_mod == 1:
    dicionario.update(gerador_trina)
    gerador_escolhido = gerador_trina # pra poder usar no cálculo N módulos
    imagem_gerador_placeholder = gerador_trina['Imagem_gerador']
elif tipo_mod == 2:
    dicionario.update(gerador_canadian)
    gerador_escolhido = gerador_canadian
    imagem_gerador_placeholder = gerador_trina['Imagem_gerador']
elif tipo_mod == 3:
    dicionario.update(gerador_astronergy) 
    gerador_escolhido = gerador_astronergy
    imagem_gerador_placeholder = gerador_trina['Imagem_gerador']
elif tipo_mod == 4: # Gokin GK-1-72HT585M (da imagem)
    dicionario.update(gerador_gokin)
    gerador_escolhido = gerador_gokin
    imagem_gerador_placeholder = gerador_trina['Imagem_gerador']
elif tipo_mod == 5: # Hanersun HN18-72H585
    dicionario.update(gerador_hanersun_585)
    gerador_escolhido = gerador_hanersun_585
    imagem_gerador_placeholder = gerador_trina['Imagem_gerador']


if tipo_inv == 1:
    dicionario.update(inversor_growatt225)
    inversor_escolhido = inversor_growatt225
    imagem_inversor_placeholder = inversor_growatt225['Imagem_inversor']
elif tipo_inv == 2:
    dicionario.update(inversor_Sungrow)
    inversor_escolhido = inversor_Sungrow
    imagem_inversor_placeholder = inversor_Sungrow['Imagem_inversor']
elif tipo_inv == 3:
    dicionario.update(inversor_Hoymiles)
    inversor_escolhido = inversor_Hoymiles
    imagem_inversor_placeholder = inversor_Hoymiles['Imagem_inversor']
elif tipo_inv == 4: # Hyxipower HYX-M2000-SW
    dicionario.update(inversor_hyxipower_m2000)
    inversor_escolhido = inversor_hyxipower_m2000
    imagem_inversor_placeholder = inversor_hyxipower_m2000['Imagem_inversor']

#Cálculo do número estimado de módulos
Pot_nom_com_virgula = float(dicionario['Pot_nom'].replace(',','.')) #porque o caralho do numero veio com ponto ao invés de vírgula
Pot_max_lista = gerador_escolhido['Pot_max'].rsplit(' ',1) #porque tem um W junto do valor
Pot_max_valor = float(Pot_max_lista[0])/1000
N_mod = round((Pot_nom_com_virgula/Pot_max_valor),2)
dicionario['N_mod'] = N_mod

#Cálculo da energia mensal produzida e total
m = 0
total_energia = 0
for key in energia_mensal.keys():
    energia_mensal[key] = round(Pot_nom_com_virgula * mes[m] * irradiacao[m] * 0.8,2)
    m += 1
    total_energia += energia_mensal[key]

dicionario.update(energia_mensal)

Total_energia = round(total_energia,2)
dicionario['Total_energia']  = Total_energia 
n = 10 
Total_arredondado = n * math.ceil(Total_energia/n)
dicionario['Total_arredondado'] = Total_arredondado


template = 'D:/UnB/Python Stuff/Memorial Descritivo - Template.docx'
output = 'Memorial Preenchido Teste.docx'

try:
    doc = DocxTemplate(template)
    dicionario['imagem_gerador'] = InlineImage(doc, imagem_gerador_placeholder, width=Cm(4.0))
    dicionario['imagem_inversor'] = InlineImage(doc, imagem_inversor_placeholder, width=Cm(3.5))

    doc.render(dicionario)
    doc.save(output)
    print(f"\nSUCESSO: Documento preenchido salvo como: {output}")

except Exception as e:
    print(f"\nERRO ao preencher o documento com Docxtpl. Verifique se o template Word está fechado e o caminho correto. Detalhes: {e}")

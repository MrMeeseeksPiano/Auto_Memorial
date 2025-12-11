import pdfplumber
import re
from docxtpl import DocxTemplate

rotulos = ['CEP','Nome','Endereco','Telefone','Email','Qtd_mod','Qtd_inv','Pot_nom']
rotulos_pdf = ['CEP da UC com GD','Nome do Titular da UC com GD','Endereço','Telefone do Titular \(DDD \+ número\)','E-mail do Titular da UC com GD','Quantidade de Módulos','Quantidade de Inversores','Potência Total dos Módulos \(kW\)']
energia_mensal = {'Energia_jan':0,'Energia_fev':0,'Energia_mar':0,'Energia_abr':0,'Energia_maio':0,'Energia_jun':0,'Energia_jul':0,'Enegia_ago':0,'Energia_set':0,'Energia_out':0,'Energia_nov':0,'Energia_dez':0}

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
                 'Peso':'38,3 kg'}

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
    'Peso': '21,8 kg'
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
                    'Nome_gerador':'Growatt',
                    'Link_gerador':'https://server.growatt.com/login',
                    'App_gerador':'ShinePhone'
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
                       'Peso_inv':'5,1 kg'}

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
    'Peso_inv': '3,75 kg'
}
dicionario = {}

print("Escolha o tipo de módulo:\n[1]Trina Solar\n[2]Canadian Solar\n[3]Astronergy")
tipo_mod = int(input(''))
print("Escolha o tipo de inversor:\n[1]Growatt\n[2]Sun Grow\n[3]Hoymiles")
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

if tipo_mod == 1:
    dicionario.update(gerador_trina)
    gerador_escolhido = gerador_trina # pra poder usar no cálculo N módulos
elif tipo_mod == 2:
    dicionario.update(gerador_canadian)
    gerador_escolhido = gerador_canadian
elif tipo_mod == 3:
    dicionario.update(gerador_astronergy) 
    gerador_escolhido = gerador_astronergy   

if tipo_inv == 1:
    dicionario.update(inversor_growatt225)
elif tipo_inv == 2:
    dicionario.update(inversor_Sungrow)
elif tipo_inv == 3:
    dicionario.update(inversor_Hoymiles)

Pot_nom_com_virgula = float(dicionario['Pot_nom'].replace(',','.')) #porque o caralho do numero veio com ponto
Pot_max_lista = gerador_escolhido['Pot_max'].rsplit(' ',1) #porque tem um W junto do valor
Pot_max_valor = float(Pot_max_lista[0])/1000
N_mod = round((Pot_nom_com_virgula/Pot_max_valor),2)
dicionario['N_mod'] = N_mod



template = 'D:/UnB/Python Stuff/Memorial Descritivo - Template.docx'
output = 'Memorial Preenchido Teste.docx'

try:
    doc = DocxTemplate(template)
    doc.render(dicionario)
    doc.save(output)
    print(f"\nSUCESSO: Documento preenchido salvo como: {output}")

except Exception as e:
    print(f"\nERRO ao preencher o documento com Docxtpl. Verifique se o template Word está fechado e o caminho correto. Detalhes: {e}")


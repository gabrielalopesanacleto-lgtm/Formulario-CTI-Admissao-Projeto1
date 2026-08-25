from time import strftime
from datetime import datetime
from openpyxl import Workbook
from openpyxl import load_workbook
import os 

def main():
    
    paciente = input("Nome do paciente: ")
    numero_atendimento = int(input("Numero do atendimento: "))
    data_de_nascimento = datetime.strptime(input("Data de nascimento (ddmmyyyy): "), "%d%m%Y")
    diagnostico = input("Diagnostico: ") 
    ims = int(input("(0 a 10) Escala IMS: "))
    nivel_de_contato = input("Nivel de contato com o paciente (alerta ou irresponsivo): ")
    tubo = input("Paciente está em uso de VMI?: ")
    oxigenio = input("Paciente está em uso de oxigenio: ")
    tabagista = input("Paciente é tabagista?: ")
    tamanho_do_tubo = "N/A"
    rima = "N/A" 
    fluxo_de_oxigenio = "N/A" 
    nome_do_acompanhante = "N/A"
    tempo_de_tabagismo = "N/A"
    quantidade_de_maços = "N/A"
    tempo_de_parada = "N/A"

    while True:
        data_atual = datetime.now()
        idade = data_atual.year - data_de_nascimento.year
        idade -= (data_atual.month, data_atual.day) < (data_de_nascimento.month, data_de_nascimento.day)
        idade = int(idade)
        if idade < 0 or idade > 150:
            print("Idade invalida, tente novamente.")
        else:
            break
        
    print("Idade do paciente: ", idade, data_de_nascimento.strftime("%d/%m/%Y"))
    print("Nome do paciente: ", paciente)
    print("Numero do atendimento: ", numero_atendimento)
    print("Diagnostico: ", diagnostico)
    
    if nivel_de_contato == "alerta":
        print("Perguntas feitas para o paciente: " + paciente)   
    elif nivel_de_contato == "irresponsivo":
        nome_do_acompanhante = input("Nome do acompanhante: ")
        print("Perguntas feitas para o acompanhante do  paciente:    " + nome_do_acompanhante)
            
    
    if oxigenio == "sim":
        fluxo_de_oxigenio = float(input("Fluxo de oxigenio: ").replace(",", "."))
        print ("Fluxo de oxigenio: ", fluxo_de_oxigenio)
    else:
        print("Paciente não está em uso de oxigenio")
            
    
    if tubo == "sim":
            tamanho_do_tubo =float(input("Tamanho do tubo: ").replace(",", "."))
            rima = float(input("Rima: ").replace(",", "."))
            print ("tamanho do tubo: ", tamanho_do_tubo, "Rima: ", rima)
    else:
            print("Paciente não está em uso de VMI")               
            
        
    if ims >=0 and ims <=4:
            print("Nivel maximo de assistencia")
    elif ims >=5 and ims <=7:
            print("Nivel moderado de assistencia.")
    elif ims >= 8 and ims <=10:
            print("Nivel minimo de assistencia.")
            
        
    if tabagista == "sim":
        tempo_de_tabagismo = int(input("Tempo de tabagismo: "))
        quantidade_de_maços = float(input("Quantidade de maços: ").replace(",", "."))
        print ("Tempo de tabagismo: ", tempo_de_tabagismo, "Quantidade de maços: ", quantidade_de_maços)
    elif tabagista == "parou":
        tempo_de_parada = float(input("Tempo de parada: ").replace(",", "."))
        print ("Tempo de parada: ", tempo_de_parada)
    else:
        print("Paciente não é tabagista")  
       

    formulario_ADM = {
        "Paciente": paciente,
        "Numero_Atendimento": numero_atendimento,
        "Idade": idade,
        "Diagnostico": diagnostico,
        "IMS": ims,
        "Nivel_de_Contato": nivel_de_contato,
        "Ventilação mecânica invasiva": tubo,
        "Tamanho do Tubo": tamanho_do_tubo,
        "Rima": rima,
        "Oxigenoterapia": oxigenio, 
        "L/min": fluxo_de_oxigenio, 
        "Tabagista": tabagista,
        "Quantidade de Macos": quantidade_de_maços,
        "Anos de Tabagismo": tempo_de_tabagismo,
        "Tempo de Parada": tempo_de_parada,      
    }   
   
    if os.path.exists ("formulario_ADM.xlsx"): 
            planilha = load_workbook("formulario_ADM.xlsx")
            aba = planilha.active
    else:
            planilha = Workbook ()
            aba = planilha.active 
            aba.append(["Nome", "Idade", "Numero de Atendimento", "Data de Nascimento", "Diagnóstico", "IMS", "Nível de Contato", "Ventilação Mecânica Invasiva", "Tamanho do tubo", "Rima", "Oxigenoterapia", "L/min", "Tabagista", "Quantidade de maços", "Anos de Tabagismo", "Tempo de Parada" ])
    aba.append(list(formulario_ADM.values()))
    planilha.save("formulario_ADM.xlsx")
main() 
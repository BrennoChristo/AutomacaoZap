import pandas as pd
import selenium
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import urllib.parse
import datetime
import threading
import json
from watchdog.events import FileSystemEventHandler

# Verifica se atualização de dado incluso na planilha
class Dados_atualizados(FileSystemEventHandler):
    def modificacao(self, dados):
        if dados.src_path == "CronogramaASO1.xlsx":
            print("Planilha modificada! Recarregando dados...")
            # Recarrega os dados da planilha
            global contatos_aso
            contatos_aso = pd.read_excel("CronogramaASO1.xlsx")

# Inicializa o navegador
navegador = webdriver.Chrome()

# Abre o site do WhatsApp Web
navegador.get("https://web.whatsapp.com/")

# Cria um tempo para o Whatsapp Web não fechar
def worker():
    # Esta thread não deve realizar operações de E/S durante o encerramento
    while True:
        time.sleep(1)


thread = threading.Thread(target=worker, daemon=True)
thread.start()

# Continua com o restante do código e aguardando a interação do usuário
while len(navegador.find_elements(By.ID, "side")) == 0:
    time.sleep(1)

# Data atual para fazer interface com a linha 50 e 80 para comunicação
dataAtual = datetime.date.today()
data_Comunicação = pd.to_datetime(dataAtual, format='%d/%m/%y')

# Arquivo de controle para colaboradores já comunicados
controle_avisos_path = "controle_avisos.json"


# Carrega os colaboradores já avisados
def carregar_controle_avisos():
    try:
        with open(controle_avisos_path, "r") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}


# Salva os colaboradores já avisados
def salvar_controle_avisos(controle_avisos):
    with open(controle_avisos_path, "w") as f:
        json.dump(controle_avisos, f)


controle_avisos = carregar_controle_avisos()

# Já estamos com login feito no WhatsApp Web
# Faz a leitura da planilha
while True:
    # Atualiza a planilha a cada iteração
    contatos_aso = pd.read_excel("CronogramaASO1.xlsx")

    for i, row in contatos_aso.iterrows():
        colaborador = row["COLABORADOR"]
        dataAso = row['DIA AGENDADO']
        horaAso = row['HORARIO']
        linkAso = row['ATENDIMENTO']
        codigoAso = row['CODIGO']
        numero = row["TELEFONE COLABORADOR"]

        # Verifica se o colaborador já foi avisado
        if colaborador in controle_avisos:
            ultima_notificacao_str = controle_avisos[colaborador]
            ultima_notificacao = datetime.datetime.strptime(ultima_notificacao_str,'%d/%m/%y')  # Adapte o formato se necessário

            # Comparando apenas as datas (ignorando a hora)
            if ultima_notificacao.date() != (dataAso - datetime.timedelta(days=1)).date():
                del controle_avisos[colaborador]
            else:
                print(f"Já foi avisado recentemente: {colaborador}")
        else:

            # Adequação do formato da Data
            dataEnvio = pd.to_datetime(dataAso, format='%d/%m/%y')

            # Verifica se falta um dia para comunicar o colaborador
            if data_Comunicação == dataEnvio - datetime.timedelta(days=1):

                # Criação da mensagem com as informações
                texto = urllib.parse.quote(f'''Seu ASO periódico é dia {dataAso.strftime('%d/%m/%y')} às {horaAso.strftime('%H:%M')}.
                Link: {linkAso}
                Código: {codigoAso}
                
                *COMO REALIZAR A CONSULTA*
                
                *Computador*
                - Só acessar o link que entrará direto.
                
                *Celular*
                - Acessar o link para baixar o app "MeuSoc";
                - Não precisa de fazer login;
                - Clicar em VideoChamada que fica no canto inferior esquerdo;
                - Colocar o código disponibilizado;
                - Entrar na reunião com a médica do trabalho.''')

                # Monta o link do WhatsApp com o número de telefone e a mensagem
                link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"
                time.sleep(5)

                # Abre o link para enviar a mensagem
                navegador.get(link)

                # Aguarda até que a página de envio do WhatsApp seja carregada
                while len(navegador.find_elements(By.ID, "side")) == 0:
                    time.sleep(5)

                # Adiciona um tempo de espera entre os envios
                time.sleep(3)  # Aguarda 3 segundos antes de enviar a próxima mensagem

                # Verificar se o número está correto
                if len(navegador.find_elements(By.XPATH,
                                               '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[1]')) == 0:

                    # Campo para aperta ENTER para enviar a mensagem
                    navegador.find_element(By.XPATH,
                                           '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/div[2]/div/p[15]/span').send_keys(
                        Keys.ENTER)
                    time.sleep(10)

                    # Salva colaboradores que foram comunicados
                    controle_avisos[colaborador] = datetime.datetime.now().strftime('%d/%m/%y')
                    salvar_controle_avisos(controle_avisos + '\n')
                else:
                    print(f'O numero inserido do {colaborador}está errado!')

            # verifica se a data do ASO já passou
            elif dataAso < data_Comunicação:
                print(f'Passou a data de avisar o colaborador {colaborador}')

            else:
                print(f'Não está na hora de comunicar o colaborador o {colaborador}')

    # Pausa o loop por um tempo antes de verificar novamente a planilha
    time.sleep(30)  # Espera 5 minutos para verificar novamente


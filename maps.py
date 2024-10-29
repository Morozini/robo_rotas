from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from openpyxl import load_workbook
import openpyxl
import pandas as pd
from tqdm import tqdm
import os
import random
from datetime import timedelta
import re
from validation_data_routes import executar,validation_tempo_total,validation_hora_data


class mapsAuto:

    def __init__(self, name_file:str):

        self.numero_aleatorio = random.randint(0, 9999)
        self.pasta = r'Z:\CELULAS\POSTAL e CASSI\2024\00 - Postal Saúde\Milena Villamil\03 - Equipe\GABRIEL\ROTAS\RELATORIO_MAPS'
        self.name_relatorio = 'Relatorio-maps'
        self.diretorio_atual = os.path.dirname(os.path.abspath(__file__))
        self.create_dir()
        self.CriandoPlanilha()
        self.get_data_rotas(name_file)

    def create_dir(self):
        if not os.path.exists(self.pasta):
            os.makedirs(self.pasta)

    def get_data_rotas(self, data):

        df = pd.read_excel(data)

        index_rotas = df[df['Name'].str.contains("Rotas de", na=False)].index
    
        origens = []
        destinos = []
        
        ultima_cidade_anterior = None
        
        for i in range(len(index_rotas)):

            start = index_rotas[i] + 1  
            end = index_rotas[i + 1] if i + 1 < len(index_rotas) else len(df)
            
            
            cidades_sequencia = df['Name'].iloc[start:end].dropna().values
            
            
            if ultima_cidade_anterior is not None and len(cidades_sequencia) > 0:
                origens.append(ultima_cidade_anterior)
                destinos.append(cidades_sequencia[0])
            
    
            for j in range(len(cidades_sequencia) - 1):
                origens.append(cidades_sequencia[j])
                destinos.append(cidades_sequencia[j + 1])
            
            
            if len(cidades_sequencia) > 0:
                ultima_cidade_anterior = cidades_sequencia[-1]

        df_pares = pd.DataFrame({
            'ORIGEM': origens,
            'DESTINO': destinos
        })

        arquivo_saida = f'.\\data\\files\\arquivo_pares_{self.numero_aleatorio}.xlsx'

        df_pares.to_excel(arquivo_saida, index=False)

        self.url_mapa(arquivo_saida)

    def url_mapa(self, data:list,data_inicial):

        url_maps = "https://www.google.com/maps?q"
     
        driver = webdriver.Chrome()

        driver.get(url_maps)

        wait = WebDriverWait(driver, 60)
        botao_rota = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="hArJGc"]')))
        botao_rota.click()

        sleep(1)

        df = pd.read_excel(data)

        tempo_total_acumulado = 0.0
        data_atual = data_inicial

        for linha in tqdm(df.index):
            tempo_atendimento_individual = 10.0
            
            origem = df.loc[linha,'ORIGEM']
            destino = df.loc[linha,'DESTINO']

            quantidade_funcionario = df.loc[linha, 'QUANTIDADE']
            
            campo_um = driver.find_element(By.XPATH, '//*[@id="sb_ifc50"]/input')
            campo_um.click()
            campo_um.clear()
            campo_um.send_keys(origem)
            
            sleep(2)

            campo_dois = driver.find_element(By.XPATH, '//*[@id="sb_ifc51"]/input')
            campo_dois.click()
            campo_dois.clear()
            campo_dois.send_keys(destino, Keys.ENTER)

            sleep(2)

            campo_img_carro = driver.find_element(By.XPATH, '//*[@id="omnibox-directions"]/div/div[2]/div/div/div/div[2]/button/img')
            campo_img_carro.click()

            sleep(2)

            kilometragem = driver.find_element(By.XPATH, '//*[@id="section-directions-trip-0"]/div[1]/div/div[1]/div[2]/div').text
            tempo_local = driver.find_element(By.XPATH, '//*[@id="section-directions-trip-0"]/div[1]/div/div[1]/div[1]').text

            tempo_local_new = self.adicionar_30_minutos(tempo_local)

            tempo_total_atual = validation_tempo_total(funcionarios=quantidade_funcionario,
                                                   tempo_atendimento=tempo_atendimento_individual,
                                                   tempo_desloc=tempo_local_new)
        
            # Validando o tempo e acumulando se necessário
            tempo_total_acumulado, data_atual = validation_hora_data(tempo_total_acumulado, 
                                                                                tempo_total_atual, 
                                                                                data_atual)

            values = [origem, destino, kilometragem, tempo_local_new]
            
            self.save_file(values)
    
    def adicionar_30_minutos(self, tempo_str):
        horas = 0
        minutos = 0
        
        match_horas = re.search(r'(\d+)\s*h', tempo_str)
        match_minutos = re.search(r'(\d+)\s*min', tempo_str)
        
        if match_horas:
            horas = int(match_horas.group(1))
        if match_minutos:
            minutos = int(match_minutos.group(1))
        
        tempo_inicial = timedelta(hours=horas, minutes=minutos)
        tempo_final = tempo_inicial + timedelta(minutes=30)
        
        total_horas = tempo_final.seconds // 3600
        total_minutos = (tempo_final.seconds % 3600) // 60
        
        return f"{total_horas}:{total_minutos}"

    def CriandoPlanilha(self):
        
        # CRIANDO A RELATORIO
        self.relatorio_excel = openpyxl.Workbook()
        self.sheet = self.relatorio_excel.active
        self.indice = ["Origem", "Destino", "KM", "TEMPO"]
        self.sheet.append(self.indice)

    def save_file(self, fields:list):
        self.INFO = fields
        for row in [self.INFO]:
            self.sheet.append(row)
            self.relatorio_excel.save(
                f'{self.pasta}\\{self.name_relatorio}_{self.numero_aleatorio}.xlsx')

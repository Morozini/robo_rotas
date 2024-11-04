from datetime import datetime, timedelta
import pandas as pd
from tqdm import tqdm
import openpyxl

class creatRoutes:

    def __init__(self, data_inicial, file, pasta):
        self.pasta = pasta
        self.name_relatorio = "relatorios_rotas"
        self.ultimo_horario = "08:00"
        self.CriandoPlanilha()

        self.executar(file, data_inicial)

    def validation_tempo_total(self, funcionarios: int, tempo_atendimento: float, tempo_desloc: float):
        tempo_atendimento_total = funcionarios * tempo_atendimento
        tempo_total = tempo_atendimento_total + tempo_desloc
        return tempo_total

    def converter_minutos_em_horas_formatadas(self, horas_float):
        total_minutos = int(horas_float)
        horas = total_minutos // 60
        minutos = total_minutos % 60
        return f"{horas:02}:{minutos:02}"

    def validation_hora_data(self, tempo_total_acumulado, tempo_atual, data_inicial_str):
        # Limite diário em minutos (7 horas = 420 minutos)
        limite_diario = 420.0
        almoco = 60.0

        # Acumular o tempo total com o tempo atual
        tempo_total_acumulado += tempo_atual

        if tempo_total_acumulado >= limite_diario:
            # Se o tempo ultrapassa o limite, zera o acumulado e ajusta a data para o próximo dia
            tempo_total_acumulado = 0.0
            data_formatada = datetime.strptime(data_inicial_str, "%d/%m/%Y")
            proxima_data = data_formatada + timedelta(days=1)
            proxima_data_str = proxima_data.strftime("%d/%m/%Y")
            return tempo_total_acumulado, proxima_data_str
        else:
            if tempo_total_acumulado >= 240.0:

                # print()
                # print(f"antes do almoco: {tempo_total_acumulado}")

                tempo_total_acumulado += almoco

                # print()
                # print(f"Depois do almoco: {tempo_total_acumulado}")

                return tempo_total_acumulado, data_inicial_str
            
            return tempo_total_acumulado, data_inicial_str

    def converter_tempo_para_minutos(self, tempo_str):
        horas, minutos = map(int, tempo_str.split(':'))
        total_minutos = horas * 60 + minutos
        return float(total_minutos)

    def executar(self, file: str, data_inicial: str):
        
        df = pd.read_excel(file)

        tempo_total_acumulado = 0.0
        data_atual = data_inicial

        for linha in tqdm(df.index):
            tempo_atendimento_individual = 10.0  # Tempo fixo por atendimento (em minutos)

            # Pegando o tempo de destino e a quantidade de funcionários
            tempo_destino = str(df.loc[linha, 'TEMPO_DESTINO'])
            quantidade_funcionario = df.loc[linha, 'QUANTIDADE']

            # Convertendo tempo de destino para minutos
            tempo_deslocamento = self.converter_tempo_para_minutos(tempo_destino)

            # Calculando o tempo total atual para a linha
            tempo_total_atual = self.validation_tempo_total(funcionarios=quantidade_funcionario,
                                                    tempo_atendimento=tempo_atendimento_individual,
                                                    tempo_desloc=tempo_deslocamento)
            
            # Validando o tempo e acumulando se necessário
            tempo_total_acumulado, data_atual = self.validation_hora_data(tempo_total_acumulado, 
                                                                                tempo_total_atual, 
                                                                                data_atual)

            tempo_novo = self.converter_minutos_em_horas_formatadas(tempo_total_acumulado)

            

            tempo_restante = self.calculo_tempo_restante(tempo_novo)

            print(f"Data Atual: {data_atual}, Tempo Total Acumulado: {tempo_novo} minutos, tempo restante: {tempo_restante}")

            resultado = self.somar_horarios(self.ultimo_horario, tempo_novo)

            self.ultimo_horario += tempo_novo

            lista_dados = [data_atual, tempo_novo, resultado]
            self.save_file(fields=lista_dados)

    def calculo_tempo_restante(self, tempo_atual):
        horario_total = timedelta(hours=8)
        tempo_acumulado_str = tempo_atual

        # Convertendo o tempo acumulado de string para timedelta
        horas, minutos = map(int, tempo_acumulado_str.split(":"))
        tempo_acumulado = timedelta(hours=horas, minutes=minutos)

        # Calculando o tempo restante
        tempo_restante = horario_total - tempo_acumulado

        # Formatando a saída
        horas_restantes = tempo_restante.seconds // 3600
        minutos_restantes = (tempo_restante.seconds // 60) % 60

        return f"{horas_restantes}:{minutos_restantes}"

    def CriandoPlanilha(self):
            
        # CRIANDO A RELATORIO
        self.relatorio_excel = openpyxl.Workbook()
        self.sheet = self.relatorio_excel.active
        self.indice = [ "DATA", "INICIO", "FIM"]
        self.sheet.append(self.indice)

    def save_file(self, fields:list):
        self.INFO = fields
        for row in [self.INFO]:
            self.sheet.append(row)
            self.relatorio_excel.save(
                f'{self.pasta}\\{self.name_relatorio}.xlsx')

    class CalculadoraTempo:
        def __init__(self):
        # o horario "00:00" sempre vai ser 08:00   
            self.ultimo_horario = "08:00"

    def somar_horarios(self, h1, h2):
            
            if ":" not in h1 or ":" not in h2:
                print("Formato de horário inválido. Use 'HH:MM'.")
                return "00:00"

            # Divide o horário e garante que há apenas horas e minutos
            h1_partes = h1.split(":")
            h2_partes = h2.split(":")
            if len(h1_partes) < 2 or len(h2_partes) < 2:
                print("Formato de horário inválido. Use 'HH:MM'.")
                return "00:00"

            # Pega apenas horas e minutos (ignora segundos, se houver)
            h1_horas, h1_minutos = int(h1_partes[0]), int(h1_partes[1])
            h2_horas, h2_minutos = int(h2_partes[0]), int(h2_partes[1])
            
            
            
            tempo1 = timedelta(hours=h1_horas, minutes=h1_minutos)
            tempo2 = timedelta(hours=h2_horas, minutes=h2_minutos)
            
            
            total_tempo = tempo1 + tempo2
            
            
            total_horas = total_tempo.seconds // 3600
            total_minutos = (total_tempo.seconds // 60) % 60

            # Retorna o resultado formatado como string
            return f"{total_horas:02}:{total_minutos:02}"
        
    def adicionar_tempo(self, tempo_novo):  

        resultado = self.somar_horarios(self.ultimo_horario, tempo_novo)
        self.ultimo_horario = resultado
        return resultado
    
if __name__ == "__main__":

    # Definir a data inicial
    data_inicial_str = "03/03/2025"

    # Caminho para o arquivo Excel
    file = r"Z:\CELULAS\POSTAL e CASSI\2024\00 - Postal Saúde\Milena Villamil\03 - Equipe\GABRIEL\ROTAS\MODELO_ROTAS_UPDATE.xlsx"
    pasta = r"C:\Users\gabriel.andrade\Desktop\convercao_new"

    creatRoutes(data_inicial_str, file, pasta)
    

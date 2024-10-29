from datetime import datetime, timedelta
import pandas as pd
from tqdm import tqdm


def validation_tempo_total(funcionarios: int, tempo_atendimento: float, tempo_desloc: float):
    tempo_atendimento_total = funcionarios * tempo_atendimento
    tempo_total = tempo_atendimento_total + tempo_desloc
    return tempo_total

def converter_minutos_em_horas_formatadas(horas_float):
    total_minutos = int(horas_float)
    horas = total_minutos // 60  # Número de horas
    minutos = total_minutos % 60  # Número de minutos restantes
    return f"{horas:02}:{minutos:02}"

def validation_hora_data(tempo_total_acumulado, tempo_atual, data_inicial_str):
    # Limite diário em minutos (7 horas = 420 minutos)
    limite_diario = 420.0

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
        return tempo_total_acumulado, data_inicial_str

# Função para converter tempo "hh:mm" em minutos
def converter_tempo_para_minutos(tempo_str):
    horas, minutos = map(int, tempo_str.split(':'))
    total_minutos = horas * 60 + minutos
    return float(total_minutos)

# Função principal para processar os dados
def executar(file: str, data_inicial: str):
    df = pd.read_excel(file)

    tempo_total_acumulado = 0.0
    data_atual = data_inicial

    for linha in tqdm(df.index):
        tempo_atendimento_individual = 10.0  # Tempo fixo por atendimento (em minutos)

        # Pegando o tempo de destino e a quantidade de funcionários
        tempo_destino = str(df.loc[linha, 'TEMPO_DESTINO'])
        quantidade_funcionario = df.loc[linha, 'QUANTIDADE']

        # Convertendo tempo de destino para minutos
        tempo_deslocamento = converter_tempo_para_minutos(tempo_destino)

        # Calculando o tempo total atual para a linha
        tempo_total_atual = validation_tempo_total(funcionarios=quantidade_funcionario,
                                                   tempo_atendimento=tempo_atendimento_individual,
                                                   tempo_desloc=tempo_deslocamento)
        
        # Validando o tempo e acumulando se necessário
        tempo_total_acumulado, data_atual = validation_hora_data(tempo_total_acumulado, 
                                                                            tempo_total_atual, 
                                                                            data_atual)
        # Exibir resultado para acompanhar o progresso
        print(f"Data Atual: {data_atual}, Tempo Total Acumulado: {tempo_total_acumulado:.2f} minutos")

if __name__ == "__main__":

    # Definir a data inicial
    data_inicial_str = "03/03/2025"

    # Caminho para o arquivo Excel
    file = r"Z:\CELULAS\POSTAL e CASSI\2024\00 - Postal Saúde\Milena Villamil\03 - Equipe\GABRIEL\ROTAS\MODELO_ROTAS_UPDATE.xlsx"

    # Executar o processo
    executar(file, data_inicial_str)

import pandas as pd

def quadroVagas():
    base = pd.read_excel("./bases/Lista geral.xlsx") # Base de dados geral.
    base['HORÁRIO ÍNICIO'] = pd.to_datetime(base['HORÁRIO ÍNICIO'], format='%H:%M:%S', errors='coerce')
    base['HORARIO FIM'] = pd.to_datetime(base['HORARIO FIM'], format='%H:%M:%S', errors='coerce')

    #Intervalos de tempo para analisar.
    quadro = {
        'HORÁRIO ÍNICIO': ['08:00:00', '10:00:00', '12:00:00', '14:00:00', '16:00:00'],
        'HORARIO FIM': ['10:00:00', '12:00:00', '14:00:00', '16:00:00', '18:00:00'],
        'SEG': [0, 0, 0, 0, 0],
        'TER': [0, 0, 0, 0, 0],
        'QUA': [0, 0, 0, 0, 0],
        'QUI': [0, 0, 0, 0, 0],
        'SEXT': [0, 0, 0, 0, 0],
        'SAB': [0, 0, 0, 0, 0],
    }

    quadro_vagas = pd.DataFrame(quadro)

    quadro_vagas['HORÁRIO ÍNICIO'] = pd.to_datetime(quadro_vagas['HORÁRIO ÍNICIO'], format='%H:%M:%S', errors='coerce')
    quadro_vagas['HORARIO FIM'] = pd.to_datetime(quadro_vagas['HORARIO FIM'], format='%H:%M:%S', errors='coerce')

    print(quadro_vagas)

    for dia in range(2, len(quadro_vagas.columns)):
        for horario in range(len(quadro_vagas[quadro_vagas.columns[dia]])):
            quadro_vagas[quadro_vagas.columns[dia]][horario] = (base.loc[(quadro_vagas.columns[dia] == base['DIA']) & (quadro_vagas['HORÁRIO ÍNICIO'][horario] >= base['HORÁRIO ÍNICIO']) & (quadro_vagas['HORARIO FIM'][horario] <= base['HORARIO FIM'])].value_counts().sum()) + (base.loc[(quadro_vagas.columns[dia] == base['DIA']) & (base['HORÁRIO ÍNICIO'] < quadro_vagas['HORARIO FIM'][horario]) & (base['HORÁRIO ÍNICIO'] > quadro_vagas['HORÁRIO ÍNICIO'][horario])].value_counts().sum()) + (base.loc[(quadro_vagas.columns[dia] == base['DIA']) & (base['HORARIO FIM'] > quadro_vagas['HORÁRIO ÍNICIO'][horario]) & (base['HORARIO FIM'] < quadro_vagas['HORARIO FIM'][horario])].value_counts().sum())
    quadro_vagas

    quadro_vagas.to_excel("./bases/QuadroVagas.xlsx", index=False, sheet_name='Assinar')
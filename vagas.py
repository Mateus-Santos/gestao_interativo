import pandas as pd

from datetime import time

from formatacao_cabecalho import alunosLista

def QuadroVagas():
    base = pd.read_excel("./bases/Lista geral.xlsx") # Base de dados geral.
    dias = ["SEG", "TER", "QUA", "QUI", "SEXT", "SAB"]

    df = pd.DataFrame({'HORÁRIO ÍNICIO': ['8:00', '10:00', '12:00', '14:00', '16:00'],
                       'horario_fim': ['10:00', '12:00', '14:00', '16:00', '18:00'],
                       'SEG':[0,0,0,0,0],
                       'TER':[0,0,0,0,0],
                       'QUA':[0,0,0,0,0],
                       'QUI':[0,0,0,0,0],
                       'SEXT':[0,0,0,0,0],
                       'SAB':[0,0,0,0,0]})

    for dia in dias:
        dia_escolhido = base[base['DIA'] == dia]

        horario1 = dia_escolhido[dia_escolhido['HORÁRIO ÍNICIO'].apply(lambda x: time(8, 0) <= x < time(10, 00))]
        vagas1 = horario1.shape[0]


        horario2 = dia_escolhido[dia_escolhido['HORÁRIO ÍNICIO'].apply(lambda x: time(10, 0) <= x < time(12, 00))]
        vagas2 = horario2.shape[0]

        horario3 = dia_escolhido[dia_escolhido['HORÁRIO ÍNICIO'].apply(lambda x: time(12, 0) <= x < time(14, 00))]
        vagas3 = horario3.shape[0]

        horario4 = dia_escolhido[dia_escolhido['HORÁRIO ÍNICIO'].apply(lambda x: time(14, 0) <= x < time(16, 00))]
        vagas4 = horario4.shape[0]

        horario5 = dia_escolhido[dia_escolhido['HORÁRIO ÍNICIO'].apply(lambda x: time(16, 0) <= x < time(18, 00))]
        vagas5 = horario5.shape[0]

        print(vagas1, vagas2, vagas3, vagas4, vagas5)

QuadroVagas()
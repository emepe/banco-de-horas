import pandas as pd
import os
from datetime import datetime
import sys
import time
import win32evtlog

evento = "Logon"
if len(sys.argv) > 1:
    evento = sys.argv[1]


arquivo = r"C:\Users\Maria.Paula\Desktop\Projeto1_AppHoras\Horários_de_entrada_x_saída.xlsx"

def registrar_evento():
    horario_atual = datetime.now()
    data_formatada = horario_atual.strftime("%d/%m/%Y")
    hora_formatada = horario_atual.strftime("%H:%M:%S")

    hora_inicio = datetime.strptime("08:40:00", "%H:%M:%S").time()
    hora_fim = datetime.strptime("12:30:00", "%H:%M:%S").time()

    novo_registro = {
        "Data": [data_formatada],
        "Hora": [hora_formatada],
        "Evento": [evento]
    }

    df_novo = pd.DataFrame(novo_registro)

    if hora_inicio <= horario_atual.time() <= hora_fim:
        print("Registro não realizado devido ao horário em que foi realizado o logon (entre 08h40 e 12h30).")

    else:
    
        if os.path.exists(arquivo):

            df_existente = pd.read_excel(arquivo)
            df_atualizado = pd.concat([df_existente, df_novo], ignore_index=True)

        else:

            df_atualizado = df_novo

        df_atualizado.to_excel(arquivo, index=False)

    print(f"Evento '{evento}' registrado com sucesso.")


def monitorar_eventos():
    log_type = "Security"
    server = "localhost"

    h_log = win32evtlog.OpenEventLog(server, log_type)

    events = win32evtlog.ReadEventLog(h_log, win32evtlog.EVENTLOG_NEWEST_FIRST, 0)

    for event in events:
        if event.EventID == 4801:
            print("Desbloqueio do computador detectado :)")
            registrar_evento()
            break

if __name__ == "__main__":
    while True:
        monitorar_eventos()
        time.sleep(5)


    #existe alguma forma de o computador saber que fiz logon e assim ele rodar o codigo com um arquivo bat? já tenho esses modelos aqui:
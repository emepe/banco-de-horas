import pandas as pd
import os
from datetime import datetime

arquivo = r"C:\Users\Maria.Paula\Desktop\Projeto1_AppHoras\Horários_de_entrada_x_saída.xlsx"

horario_atual = datetime.now()
data_formatada = horario_atual.strftime("%d/%m/%Y")
hora_formatada = horario_atual.strftime("%H:%M:%S")

novo_registro = {
    "Data": [data_formatada],
    "Hora": [hora_formatada],
    "Evento": ["Logoff"]
}

df_novo = pd.DataFrame(novo_registro)

if os.path.exists(arquivo):

        df_existente = pd.read_excel(arquivo)
        df_atualizado = pd.concat([df_existente, df_novo], ignore_index=True)

else:

        df_atualizado = df_novo

df_atualizado.to_excel(arquivo, index=False)

print(f"Evento 'Logoff' registrado com sucesso.")
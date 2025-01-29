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
    "Evento": ["Desligado"]
}

df_novo = pd.DataFrame(novo_registro)

if os.path.exists(arquivo):

        df_existente = pd.read_excel(arquivo)
        df_atualizado = pd.concat([df_existente, df_novo], ignore_index=True)

else:

        df_atualizado = df_novo

df_atualizado.to_excel(arquivo, index=False)

print(f"Evento 'Desligado' registrado com sucesso.")

# Aqui tenho um código que cria uma planilha todas as vezes que o código é executado. Além disso o código insere dentro da planilha o dia e horário em que foi rodado.

# O objetivo do código é inserir o dia e horário em que o computador foi ligado e desligado dentro dessa planilha.

# Nos próximos passos eu tenho 2 opções:

    # 1) Ajustar o código corretamente (para que ele não crie uma planilha e sim adicione informações e colocar cada dia em uma linha diferente)
    # 2) Configurar o código para ser rodado sempre que ligado ou desligado o computador #
import pandas as pd
import datetime

planilha = r'C:\Users\moesios\Desktop\Final\distribuição.xlsx'
df_viagens = pd.read_excel(planilha, sheet_name='Planilha1')

if df_viagens['Hora_inicio'].dtype == object:
    try:
        df_viagens['Hora_inicio'] = pd.to_datetime(df_viagens['Hora_inicio'], format='%H:%M:%S').dt.time
    except ValueError:
        raise ValueError("'Hora_inicio' deve ser 'HH:MM:SS' ")

if pd.api.types.is_datetime64_any_dtype(df_viagens['Hora_inicio']):
    df_viagens['Hora_inicio_time'] = df_viagens['Hora_inicio'].dt.time
elif isinstance(df_viagens['Hora_inicio'][0], datetime.time):
    df_viagens['Hora_inicio_time'] = df_viagens['Hora_inicio']
else:
    raise ValueError("'Hora_inicio' deve ou datetime ou type time")

df_viagens['Hora_inicio_datetime'] = [
    datetime.datetime.combine(datetime.date.today(), time) for time in df_viagens['Hora_inicio_time']
]

intervalos = pd.date_range(start='00:00', end='23:59', freq='1H').strftime('%H:%M').tolist()
intervalos = [f"{intervalos[i]}-{intervalos[i+1]}" for i in range(len(intervalos)-1)]
intervalos.append('23:00-00:00')

labels = [f"{str(i).zfill(2)}:00-{str(i+1).zfill(2)}:00" for i in range(23)]
labels.append('23:00-00:00')

df_viagens['Intervalo'] = pd.cut(df_viagens['Hora_inicio_datetime'].dt.hour, bins=range(25), right=False, labels=labels)

grouped = df_viagens.groupby(['Intervalo', 'MOTIVO VIAGEM']).size().unstack(fill_value=0)

total_por_intervalo = grouped.sum(axis=1)

grouped_porcentagem = grouped.div(total_por_intervalo, axis=0) * 100

grouped.to_excel(r'C:\Users\moesios\Desktop\Final\deslocamentos_por_horaBO.xlsx')

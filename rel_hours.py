from icalendar import Calendar

import pandas as pd

from datetime import datetime, time, timedelta

from dateutil import rrule

import os

 

# --- CONFIGURAÇÃO DO EXPEDIENTE ---

HORA_INICIO_TRABALHO = time(8, 0)

HORA_FIM_TRABALHO = time(17, 0)

ALMOCO_INICIO = time(12, 0)

ALMOCO_FIM = time(13, 0)

 

# --- AJUSTE DE DATA ---

data_foco = datetime(2026, 5, 4)

data_inicio_busca = datetime.combine(data_foco, time.min)

data_fim_busca = datetime.combine(data_foco + timedelta(days=6), time.max)

 

if not os.path.exists('agenda0405.ics'):

    print("Erro: Arquivo 'agenda0405.ics' não encontrado na pasta.")

    exit()

 

with open('agenda0405.ics', 'rb') as f:

    gcal = Calendar.from_ical(f.read())

 

# --- FUNÇÃO DE LIMPEZA DE TEXTO (EVITA "none") ---

def limpa_texto(x):

    if x is None:

        return ''

    s = str(x).strip()

    if s.lower() == "none":

        return ''

    return s

 

 

eventos_raw = []

 

# ----------- EXTRAÇÃO CORRIGIDA DOS EVENTOS -----------

for component in gcal.walk():

 

    if component.name != "VEVENT":

        continue

 

    # Ignorar eventos cancelados

    status = component.get('status')

    if status and str(status).strip().upper() == "CANCELLED":

        continue

 

    # Captura segura do título

    summary = limpa_texto(component.get('summary'))

    description = limpa_texto(component.get('description'))

 

    # Se não houver summary nem description → ignorar evento completamente

    if not summary and not description:

        continue

 

    resumo = summary if summary else description

 

    # Captura segura das datas

    dtstart = component.get('dtstart')

    dtend = component.get('dtend')

    if not dtstart or not dtend:

        continue

 

    dtstart = dtstart.dt

    dtend = dtend.dt

 

    if not isinstance(dtstart, datetime):

        dtstart = datetime.combine(dtstart, time.min)

    if not isinstance(dtend, datetime):

        dtend = datetime.combine(dtend, time.min)

 

    dtstart = dtstart.replace(tzinfo=None)

    dtend = dtend.replace(tzinfo=None)

 

    # Recorrência

    rrule_field = component.get('rrule')

    if rrule_field:

        try:

            regra = rrule.rrulestr(rrule_field.to_ical().decode('utf-8'), dtstart=dtstart)

            for oco in regra.between(data_inicio_busca, data_fim_busca, inc=True):

                eventos_raw.append({

                    'inicio': oco,

                    'fim': oco + (dtend - dtstart),

                    'desc': resumo

                })

        except:

            if data_inicio_busca <= dtstart <= data_fim_busca:

                eventos_raw.append({'inicio': dtstart, 'fim': dtend, 'desc': resumo})

    else:

        if data_inicio_busca <= dtstart <= data_fim_busca:

            eventos_raw.append({'inicio': dtstart, 'fim': dtend, 'desc': resumo})

 

 

# --- SE NÃO HOUVER EVENTOS ---

if not eventos_raw:

    print("Nenhum evento encontrado na agenda.")

    df_raw = pd.DataFrame(columns=['inicio', 'fim', 'desc'])

else:

    df_raw = pd.DataFrame(eventos_raw).sort_values('inicio')

 

 

# ----------- GERAÇÃO DAS LINHAS PARA A PLANILHA -----------

dados_finais = []

 

for i in range(5): # Segunda a Sexta

    dia_atual = (data_foco + timedelta(days=i)).date()

 

    expediente_inicio = datetime.combine(dia_atual, HORA_INICIO_TRABALHO)

    expediente_fim = datetime.combine(dia_atual, HORA_FIM_TRABALHO)

    almoco_ini = datetime.combine(dia_atual, ALMOCO_INICIO)

    almoco_fim = datetime.combine(dia_atual, ALMOCO_FIM)

 

    eventos_dia = df_raw[df_raw["inicio"].dt.date == dia_atual].to_dict("records") if not df_raw.empty else []

 

    ponteiro = expediente_inicio

 

    for ev in eventos_dia:

        if ev['inicio'] > ponteiro:

            if ponteiro < almoco_ini and ev['inicio'] > almoco_fim:

                dados_finais.append({'Data': dia_atual, 'Inicio': ponteiro, 'Fim': almoco_ini, 'Descrição': 'JIRA'})

                dados_finais.append({'Data': dia_atual, 'Inicio': almoco_fim, 'Fim': ev['inicio'], 'Descrição': 'JIRA'})

            elif not (ponteiro >= almoco_ini and ponteiro < almoco_fim):

                dados_finais.append({'Data': dia_atual, 'Inicio': ponteiro, 'Fim': min(ev['inicio'], almoco_ini if ponteiro < almoco_ini else ev['inicio']), 'Descrição': 'JIRA'})

 

        dados_finais.append({'Data': dia_atual, 'Inicio': ev['inicio'], 'Fim': ev['fim'], 'Descrição': ev['desc']})

        ponteiro = max(ponteiro, ev['fim'])

 

    # Final do dia

    if ponteiro < expediente_fim:

        if ponteiro < almoco_ini:

            dados_finais.append({'Data': dia_atual, 'Inicio': ponteiro, 'Fim': almoco_ini, 'Descrição': 'JIRA'})

            dados_finais.append({'Data': dia_atual, 'Inicio': almoco_fim, 'Fim': expediente_fim, 'Descrição': 'JIRA'})

        elif ponteiro >= almoco_fim:

            dados_finais.append({'Data': dia_atual, 'Inicio': ponteiro, 'Fim': expediente_fim, 'Descrição': 'JIRA'})

 

 

# ----------- EXPORTAR EXCEL -----------

if dados_finais:

    final_list = []

    for d in dados_finais:

        if d['Inicio'] >= d['Fim']:

            continue

        dur = d['Fim'] - d['Inicio']

        h, m = divmod(int(dur.total_seconds()) // 60, 60)

        final_list.append({

            'Data': d['Data'].strftime('%d/%m/%Y'),

            'Hora (Início)': d['Inicio'].strftime('%H:%M'),

            'Hora (Fim)': d['Fim'].strftime('%H:%M'),

            'Descrição': d['Descrição'],

            'Total': f"{h}:{m:02d}",

            '_sort': d['Inicio']

        })

 

    df_final = pd.DataFrame(final_list).sort_values('_sort').drop(columns=['_sort'])

    df_final.to_excel('Relatorio_JIRA_0405.xlsx', index=False)

    print("Relatório gerado com sucesso!")

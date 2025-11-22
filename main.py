import pandas as pd
import gspread
import requests
import time
import base64
from datetime import datetime, timedelta
from pytz import timezone
import os
import json

# --- CONSTANTES ---
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
NOME_ABA = 'Base Pending Tratado'
INTERVALO = 'A:F'
CAMINHO_IMAGEM = "alerta.gif"
FUSO_HORARIO_SP = timezone('America/Sao_Paulo')

# 👥 DICIONÁRIO DE PESSOAS POR TURNO (COM IDS REAIS!)
TURNO_PARA_IDS = {
    "Turno 1": [
        "1461929762",  # Iromar Souza
        "1449480651",  # Ana Julia Lopes
        "9465967606",  # Fidel Lúcio
        "1268695707"   # Claudio Olivatto
    ],
    "Turno 2": [
        "9356934188",  # Fabrício Damasceno
        "1160266193",  # Joao Pedro Araujo Silva
        "1386559133",  # Murilo Santana
        "1298055860",  # Matheus Damas
        "1432898616"   # Leonardo Caus
        
    ],
    "Turno 3": [
        "9289770437",  # Fernando Aparecido da Costa
        "1436962469",  # Jose Guilherme Paco
        "9474534910",  # Kaio Baldo
        "1499919880"   # Sandor Nemes
    ]
}

# 🗓️ CONFIGURAÇÃO DE FOLGAS (0=Segunda ... 6=Domingo)
# 0=Seg, 1=Ter, 2=Qua, 3=Qui, 4=Sex, 5=Sab, 6=Dom
DIAS_DE_FOLGA = {
    # --- Turno 1 ---
    "1461929762": [5, 6],    # Iromar Souza (Dom)
    "1449480651": [5, 6], # Ana Julia Lopes (Sab, Dom)
    "9465967606": [5, 6], # Fidel Lúcio (Sab, Dom)
    "1268695707": [6],    # Claudio Olivatto (Dom)

    # --- Turno 2 ---
    "9356934188": [5, 6],   # Fabrício Damasceno (Sab, Dom)
    "1160266193": [6],      # Joao Pedro Araujo (Dom)
    "1386559133": [6, 0],   # Murilo Santana (Dom, Seg)
    "1298055860": [6],      # Matheus Damas (Dom)
    "1432898616": [0, 1, 2, 3, 4, 5], # Leonardo Caus (Folga Seg a Sab)

    # --- Turno 3 ---
    "9289770437": [6, 0],      # Fernando Aparecido (Dom, Seg)
    "1436962469": [6, 0],      # Jose Guilherme Paco (Dom, Seg)
    "9474534910": [6, 0],      # Kaio Baldo (Dom, Seg)
    "1499919880": [6]          # Sandor Nemes (Dom)
}

def identificar_turno_atual(agora):
    """Identifica o turno atual baseado na hora de São Paulo."""
    hora = agora.hour

    if 6 <= hora < 14:
        return "Turno 1"
    elif 14 <= hora < 22:
        return "Turno 2"
    else:
        return "Turno 3"

def filtrar_quem_esta_de_folga(ids_do_turno, agora, turno_atual):
    """Remove da lista de IDs quem tem folga no dia da semana atual."""
    
    # --- LÓGICA DE MADRUGADA (T3) ---
    # Se for T3 e estiver entre 00h e 06h, o turno pertence ao dia anterior.
    data_referencia = agora
    if turno_atual == "Turno 3" and agora.hour < 6:
        data_referencia = agora - timedelta(days=1)
        print("🌙 Madrugada T3: Verificando escala baseada no dia anterior (Início do Turno).")
    
    dia_semana_referencia = data_referencia.weekday() # 0=Segunda ... 6=Domingo
    
    ids_validos = []
    nomes_dias = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
    print(f"📅 Dia de referência para escala: {nomes_dias[dia_semana_referencia]}.")

    for uid in ids_do_turno:
        dias_off_da_pessoa = DIAS_DE_FOLGA.get(uid, [])
        
        if dia_semana_referencia in dias_off_da_pessoa:
            print(f"🏖️ ID {uid} está de folga (ref: {nomes_dias[dia_semana_referencia]}). Não será marcado.")
        else:
            ids_validos.append(uid)
            
    return ids_validos

def autenticar_google():
    creds_json_str = os.environ.get('GOOGLE_SERVICE_ACCOUNT_JSON')
    if not creds_json_str:
        print("❌ Erro: Variável de ambiente 'GOOGLE_SERVICE_ACCOUNT_JSON' não definida.")
        return None

    try:
        creds_dict = json.loads(creds_json_str)
        cliente = gspread.service_account_from_dict(creds_dict, scopes=SCOPES)
        print("✅ Cliente autenticado.")
        return cliente
    except Exception as e:
        print(f"❌ Erro ao autenticar: {e}")
        return None

def formatar_doca(doca):
    doca = str(doca).strip() # Garante string
    if not doca or doca == '-':
        return "Doca --"
    elif doca.startswith("EXT.OUT"):
        numeros = ''.join(filter(str.isdigit, doca))
        return f"Doca {numeros}"
    elif not doca.startswith("Doca"):
        return f"Doca {doca}"
    else:
        return doca

def obter_dados_expedicao(cliente, spreadsheet_id):
    if not cliente:
        return None, "⚠️ Cliente não autenticado."

    try:
        planilha = cliente.open_by_key(spreadsheet_id)
        aba = planilha.worksheet(NOME_ABA)
        dados = aba.get(INTERVALO)
    except Exception as e:
        return None, f"⚠️ Erro ao acessar planilha: {e}"

    if not dados or len(dados) < 2:
        return None, "⚠️ Nenhum dado encontrado."

    df = pd.DataFrame(dados[1:], columns=dados[0])
    df.columns = df.columns.str.strip()

    for col in ['Doca', 'LH Trip Number', 'Station Name', 'CPT']:
        if col not in df.columns:
            return None, f"⚠️ Coluna '{col}' não encontrada."

    df = df[df['LH Trip Number'].str.strip() != '']
    # Converte CPT inicialmente, mas a lógica real de tempo está no montar_mensagem
    df['CPT'] = pd.to_datetime(df['CPT'], dayfirst=True, errors='coerce')
    df = df.dropna(subset=['CPT'])

    return df, None

def montar_mensagem_alerta(df):
    agora = datetime.now(FUSO_HORARIO_SP)

    df = df.copy()
    # Garante datetime e trata timezone
    df['CPT'] = pd.to_datetime(df['CPT'], dayfirst=True, errors='coerce')
    df = df.dropna(subset=['CPT'])
    
    # Se o CPT da planilha não tiver fuso, localiza como SP. Se tiver, converte.
    df['CPT'] = df['CPT'].dt.tz_localize(FUSO_HORARIO_SP, ambiguous='NaT', nonexistent='NaT')
    
    df = df.dropna(subset=['CPT'])
    df['minutos_restantes'] = ((df['CPT'] - agora).dt.total_seconds() // 60).astype(int)
    df = df[df['minutos_restantes'] >= 0]

    # --- ORDENAÇÃO: Mais urgentes (menos tempo) primeiro ---
    df = df.sort_values(by='minutos_restantes', ascending=True)

    def agrupar_minutos(minutos):
        if 21 <= minutos <= 30: return 30
        elif 11 <= minutos <= 20: return 20
        elif 1 <= minutos <= 10: return 10
        else: return None

    df['grupo_alerta'] = df['minutos_restantes'].apply(agrupar_minutos)
    df_filtrado = df.dropna(subset=['grupo_alerta'])

    if df_filtrado.empty:
        return None

    mensagens = []
    # --- LOOP ORDENADO: 10 (urgente), depois 20, depois 30 ---
    for minuto in [10, 20, 30]:
        grupo = df_filtrado[df_filtrado['grupo_alerta'] == minuto]
        if not grupo.empty:
            mensagens.append("")
            mensagens.append(f"⚠️ Atenção, LTs próximas do CPT! (Faixa {minuto} min) ⚠️")
            mensagens.append("") 
            
            for _, row in grupo.iterrows():
                lt = row['LH Trip Number'].strip()
                destino = row['Station Name'].strip()
                doca = formatar_doca(row['Doca'])
                cpt_str = row['CPT'].strftime('%H:%M') 
                minutos_reais = int(row['minutos_restantes'])
                
                mensagens.append(f"🚛 {lt}")
                mensagens.append(f"{doca}")
                mensagens.append(f"Destino: {destino}")
                mensagens.append(f"CPT: {cpt_str} (faltam {minutos_reais} min)")
                mensagens.append("") 

    if mensagens and mensagens[-1] == "":
        mensagens.pop()

    return "\n".join(mensagens)

def enviar_imagem(webhook_url: str, caminho_imagem: str = CAMINHO_IMAGEM):
    if not webhook_url:
        print("❌ WEBHOOK_URL não definida.")
        return False
    try:
        if not os.path.exists(caminho_imagem):
            print(f"⚠️ Aviso: Imagem '{caminho_imagem}' não encontrada localmente. Pulando envio de imagem.")
            return False

        with open(caminho_imagem, "rb") as f:
            raw_image_content = f.read()
            base64_encoded_image = base64.b64encode(raw_image_content).decode("utf-8")
        payload = {"tag": "image", "image_base64": {"content": base64_encoded_image}}
        response = requests.post(webhook_url, json=payload)
        response.raise_for_status()
        print("✅ Imagem enviada com sucesso.")
        return True
    except Exception as e:
        print(f"❌ Erro ao enviar imagem: {e}")
        return False

def enviar_webhook_com_mencao_oficial(mensagem_texto: str, webhook_url: str, user_ids: list = None):
    if not webhook_url:
        print("❌ WEBHOOK_URL não definida.")
        return

    mensagem_final = f"{mensagem_texto}"
    payload = {
        "tag": "text",
        "text": { "format": 1, "content": mensagem_final }
    }

    if user_ids:
        user_ids_validos = [uid for uid in user_ids if uid and uid.strip()]
        if user_ids_validos:
            payload["text"]["mentioned_list"] = user_ids_validos
            print(f"✅ Enviando menção para: {user_ids_validos}")
        else:
            print("⚠️ Nenhum ID válido para marcar.")

    try:
        response = requests.post(webhook_url, json=payload)
        response.raise_for_status()
        print("✅ Mensagem com menção OFICIAL enviada com sucesso.")
    except Exception as e:
        print(f"❌ Falha ao enviar mensagem: {e}")

def main():
    webhook_url = os.environ.get('SEATALK_WEBHOOK_URL')
    spreadsheet_id = os.environ.get('SPREADSHEET_ID')

    if not webhook_url or not spreadsheet_id:
        print("❌ Variáveis SEATALK_WEBHOOK_URL ou SPREADSHEET_ID não definidas.")
        return

    cliente = autenticar_google()
    if not cliente:
        return

    # Define 'agora' UMA VEZ aqui para usar no turno e no filtro de folga
    agora = datetime.now(FUSO_HORARIO_SP)

    df, erro = obter_dados_expedicao(cliente, spreadsheet_id)
    if erro:
        print(erro)
        return

    mensagem = montar_mensagem_alerta(df)

    if mensagem:
        turno_atual = identificar_turno_atual(agora) 
        ids_brutos = TURNO_PARA_IDS.get(turno_atual, [])

        # --- APLICA FILTRO DE FOLGAS COM LÓGICA T3 MADRUGADA ---
        ids_para_marcar = filtrar_quem_esta_de_folga(ids_brutos, agora, turno_atual)
        # -------------------------------------------------------

        print(f"🕒 Turno atual: {turno_atual}")
        print(f"👥 IDs originais: {len(ids_brutos)} | IDs após filtro: {len(ids_para_marcar)}")

        enviar_webhook_com_mencao_oficial(mensagem, webhook_url, user_ids=ids_para_marcar)
        enviar_imagem(webhook_url)
    else:
        print("✅ Nenhuma LT nos critérios de alerta. Nada enviado.")

if __name__ == "__main__":
    main()

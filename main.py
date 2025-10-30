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

# 👥 DICIONÁRIO DE PESSOAS POR TURNO (COM IDS REAIS!)
TURNO_PARA_IDS = {
    "Turno 1": [
        "1461929762",  # Iromar Souza
        "9465967606",  # Fidel Lúcio
        "1268695707"   # Claudio Olivatto
    ],
    "Turno 2": [
        "9356934188",  # Fabrício Damasceno
        "1386559133",  # Murilo Santana
        "1298055860"   # Matheus Damas
    ],
    "Turno 3": [
        "9289770437",  # Fernando Aparecido da Costa
        "9474534910",  # Kaio Baldo
        "1499919880"   # Sandor Nemes
    ]
}


def identificar_turno_atual():
    """Identifica o turno atual baseado na hora de São Paulo."""
    agora = datetime.now(timezone('America/Sao_Paulo'))
    hora = agora.hour

    if 6 <= hora < 14:
        return "Turno 1"
    elif 14 <= hora < 22:
        return "Turno 2"
    else:
        return "Turno 3"


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
    doca = doca.strip()
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
    df['CPT'] = pd.to_datetime(df['CPT'], dayfirst=True, errors='coerce')
    df = df.dropna(subset=['CPT'])

    return df, None


def montar_mensagem_alerta(df):
    tz = timezone('America/Sao_Paulo')
    agora = datetime.now(tz)

    df = df.copy()
    df['CPT'] = pd.to_datetime(df['CPT'], dayfirst=True, errors='coerce')
    df = df.dropna(subset=['CPT'])
    df['CPT'] = df['CPT'].dt.tz_localize(tz, ambiguous='NaT', nonexistent='NaT')
    df = df.dropna(subset=['CPT'])
    df['minutos_restantes'] = ((df['CPT'] - agora).dt.total_seconds() // 60).astype(int)
    df = df[df['minutos_restantes'] >= 0]

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
    for minuto in [30, 20, 10]:
        grupo = df_filtrado[df_filtrado['grupo_alerta'] == minuto]
        if not grupo.empty:
            
            mensagens.append(f"⚠️ Atenção!!!")
            
            # ✨ ALTERAÇÃO (1/2): Adicionando DUAS linhas em branco
            mensagens.append("") 
            mensagens.append("") 
            
            for _, row in grupo.iterrows():
                lt = row['LH Trip Number'].strip()
                destino = row['Station Name'].strip()
                doca = formatar_doca(row['Doca'])
                cpt_str = row['CPT'].strftime('%H:%M') 
                minutos_reais = int(row['minutos_restantes'])
                
                # Formato em 4 linhas
                mensagens.append(f"🚛 {lt}")
                mensagens.append(f"{doca}")
                mensagens.append(f"Destino: {destino}")
                mensagens.append(f"CPT: {cpt_str} (faltam {minutos_reais} min)")
                
                # Linha em branco antes da próxima LT
                mensagens.append("") 

    if mensagens and mensagens[-1] == "":
        mensagens.pop()

    return "\n".join(mensagens)


def enviar_imagem(webhook_url: str, caminho_imagem: str = CAMINHO_IMAGEM):
    if not webhook_url:
        print("❌ WEBHOOK_URL não definida.")
        return False
    try:
        with open(caminho_imagem, "rb") as f:
            raw_image_content = f.read()
            base64_encoded_image = base64.b64encode(raw_image_content).decode("utf-8")
        payload = {"tag": "image", "image_base64": {"content": base64_encoded_image}}
        response = requests.post(webhook_url, json=payload)
        response.raise_for_status()
        print("✅ Imagem enviada com sucesso.")
        return True
    except FileNotFoundError:
        print(f"❌ Arquivo '{caminho_imagem}' não encontrado. Pulando imagem...")
        return False
    except Exception as e:
        print(f"❌ Erro ao enviar imagem: {e}")
        return False


def enviar_webhook_com_mencao_oficial(mensagem_texto: str, webhook_url: str, user_ids: list = None):
    if not webhook_url:
        print("❌ WEBHOOK_URL não definida.")
        return

    # A mensagem final é apenas o corpo do alerta
    mensagem_final = f"{mensagem_texto}"

    payload = {
        "tag": "text",
        "text": {
            "format": 1,
            "content": mensagem_final
        }
    }

    # ✨ ALTERAÇÃO (2/2): O bloco 'mentioned_list' foi DESATIVADO.
    # Isso remove o PING e os nomes automáticos no topo,
    # garantindo o formato limpo que você pediu.
    if user_ids:
        user_ids_validos = [uid for uid in user_ids if uid and uid.strip()]
        if user_ids_validos:
            
            # A linha abaixo foi COMENTADA para parar o ping:
            # payload["text"]["mentioned_list"] = user_ids_validos
            
            print(f"✅ Mensagem será enviada SEM menção (silenciosa).")
            print(f"   (IDs que seriam marcados: {user_ids_validos})")
        else:
            print("⚠️ Nenhum ID válido para marcar.")

    try:
        response = requests.post(webhook_url, json=payload)
        response.raise_for_status()
        print("✅ Mensagem (silenciosa) enviada com sucesso.")
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

    df, erro = obter_dados_expedicao(cliente, spreadsheet_id)
    if erro:
        print(erro)
        return

    mensagem = montar_mensagem_alerta(df)

    if mensagem:
        turno_atual = identificar_turno_atual()
        ids_para_marcar = TURNO_PARA_IDS.get(turno_atual, [])

        print(f"🕒 Turno atual: {turno_atual}")
        print(f"👥 IDs configurados para este turno: {ids_para_marcar}")

        enviar_imagem(webhook_url)
        # Os IDs são passados para a função, mas ela agora está configurada
        # para não usá-los para notificação (ping).
        enviar_webhook_com_mencao_oficial(mensagem, webhook_url, user_ids=ids_para_marcar)
    else:
        print("✅ Nenhuma LT nos critérios de alerta. Nada enviado.")


if __name__ == "__main__":
    main()

import pandas as pd
import gspread
import requests
import time
from datetime import datetime, timedelta
from pytz import timezone
import os
import json  # Para carregar o JSON das credenciais diretamente

# --- CONSTANTES GLOBAIS ---
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
NOME_ABA = 'Base Pending Tratado'
INTERVALO = 'A:F'

# --- AUTENTICAÇÃO CORRIGIDA (SEM BASE64!) ---
def autenticar_google():
    """Autentica usando o Secret JSON do GitHub (NÃO é base64!)."""

    # 1. Lê o JSON puro do segredo
    creds_json_str = os.environ.get('GOOGLE_SERVICE_ACCOUNT_JSON')
    if not creds_json_str:
        print("❌ Erro: Variável de ambiente 'GOOGLE_SERVICE_ACCOUNT_JSON' não definida.")
        return None

    try:
        # 2. Carrega a string JSON diretamente em um dicionário — SEM base64!
        creds_dict = json.loads(creds_json_str)

        # 3. Autentica com o dicionário
        cliente = gspread.service_account_from_dict(creds_dict, scopes=SCOPES)
        print("✅ Cliente gspread autenticado com Service Account (via JSON direto).")
        return cliente

    except json.JSONDecodeError as e:
        print(f"❌ Erro ao decodificar o JSON das credenciais: {e}")
        print("   Isso geralmente significa que o segredo está corrompido ou foi salvo incorretamente.")
        return None
    except Exception as e:
        print(f"❌ Erro ao autenticar com Service Account: {e}")
        return None


def identificar_turno(hora):
    if 6 <= hora < 14:
        return "Turno 1"
    elif 14 <= hora < 22:
        return "Turno 2"
    else:
        return "Turno 3"


def obter_dados_expedicao(cliente, spreadsheet_id):
    if not cliente:
        return None, "⚠️ Não foi possível autenticar o cliente."

    try:
        planilha = cliente.open_by_key(spreadsheet_id)
        aba = planilha.worksheet(NOME_ABA)
        dados = aba.get(INTERVALO)
    except Exception as e:
        return None, f"⚠️ Erro ao acessar planilha: {e}"

    if not dados or len(dados) < 2:
        return None, "⚠️ Nenhum dado encontrado na planilha."

    df = pd.DataFrame(dados[1:], columns=dados[0])
    df.columns = df.columns.str.strip()

    for col in ['Doca', 'LH Trip Number', 'Station Name', 'CPT']:
        if col not in df.columns:
            return None, f"⚠️ Coluna '{col}' não encontrada."

    df = df[df['LH Trip Number'].str.strip() != '']
    df['CPT'] = pd.to_datetime(df['CPT'], dayfirst=True, errors='coerce')
    df = df.dropna(subset=['CPT'])
    df['Turno'] = df['CPT'].dt.hour.apply(identificar_turno)

    return df, None


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


def montar_mensagem(df):
    agora = datetime.now(timezone('America/Sao_Paulo')).replace(tzinfo=None)
    limite_2h = agora + timedelta(hours=2)
    turno_atual = identificar_turno(agora.hour)

    mensagens = []
    totais = df['Turno'].value_counts().to_dict()

    df_2h = df[(df['CPT'] >= agora) & (df['CPT'] < limite_2h)].copy()
    if df_2h.empty:
        mensagens.append("🚛 LTs pendentes:\n\n✅ Sem LT pendente para as próximas 2h.\n")
    else:
        mensagens.append("🚛 LTs pendentes:\n")
        df_2h['Hora'] = df_2h['CPT'].dt.hour

        for hora, grupo in df_2h.groupby('Hora', sort=True):
            qtd_lhs = len(grupo)
            mensagens.append(f"{qtd_lhs} LH{'s' if qtd_lhs > 1 else ''} pendente{'s' if qtd_lhs > 1 else ''} às {hora:02d}h\n")
            for _, row in grupo.iterrows():
                lt = row['LH Trip Number'].strip()
                destino = row['Station Name'].strip()
                cpt = row['CPT']
                cpt_str = cpt.strftime('%H:%M')
                doca = formatar_doca(row['Doca'])

                minutos = int((cpt - agora).total_seconds() // 60)
                if minutos < 0:
                    prefixo = "❗️"
                    status = "(ATRASADO)"
                elif minutos <= 10:
                    prefixo = "⚠️"
                    status = f"(FALTAM {minutos} MIN)"
                else:
                    prefixo = ""
                    status = ""

                mensagens.append(f"{prefixo} {lt} | {doca} | Destino: {destino} | CPT: {cpt_str} {status}".strip())
            mensagens.append("")

    mensagens.append("─" * 40)
    mensagens.append("LH´s pendentes para os próximos turnos:\n")

    prioridades_turno = {
        'Turno 1': ['Turno 2', 'Turno 3'],
        'Turno 2': ['Turno 3', 'Turno 1'],
        'Turno 3': ['Turno 1', 'Turno 2']
    }

    for turno in prioridades_turno.get(turno_atual, []):
        qtd = totais.get(turno, 0)
        mensagens.append(f"⚠️ {qtd} LH{'s' if qtd != 1 else ''} pendente{'s' if qtd != 1 else ''} no {turno}")

    return "\n".join(mensagens)


def enviar_webhook(mensagem, webhook_url):
    if not webhook_url:
        print("❌ Erro: WEBHOOK_URL não fornecida.")
        return
    try:
        payload = {
            "tag": "text",
            "text": {
                "format": 1,
                "content": mensagem
            }
        }
        response = requests.post(webhook_url, json=payload)
        response.raise_for_status()
        print("✅ Mensagem enviada com sucesso.")
    except Exception as e:
        print(f"❌ Erro ao enviar mensagem: {e}")


def enviar_em_blocos(mensagem, webhook_url, limite=3000):
    linhas = mensagem.split('\n')
    bloco = []
    for linha in linhas:
        bloco.append(linha)
        if len("\n".join(bloco)) > limite:
            bloco.pop()
            enviar_webhook("```\n" + "\n".join(bloco) + "\n```", webhook_url)
            time.sleep(1)
            bloco = [linha]
    if bloco:
        enviar_webhook("```\n" + "\n".join(bloco) + "\n```", webhook_url)


def main():
    webhook_url = os.environ.get('SEATALK_WEBHOOK_URL')
    spreadsheet_id = os.environ.get('SPREADSHEET_ID')

    if not webhook_url or not spreadsheet_id:
        print("❌ Erro: Variáveis de ambiente SEATALK_WEBHOOK_URL e/ou SPREADSHEET_ID não definidas.")
        return

    cliente = autenticar_google()  # Agora lê GOOGLE_SERVICE_ACCOUNT_JSON
    if not cliente:
        print("❌ Falha na autenticação. Encerrando.")
        return

    df, erro = obter_dados_expedicao(cliente, spreadsheet_id)
    if erro:
        print(erro)
        return

    mensagem = montar_mensagem(df)
    enviar_em_blocos(mensagem, webhook_url)


if __name__ == "__main__":
    main()

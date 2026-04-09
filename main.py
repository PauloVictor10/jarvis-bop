"""
Agente Inteligente — Backend FastAPI
Suporta: OpenAI, Anthropic Claude, Ollama
TTS: OpenAI Text-to-Speech
Google Sheets: Registro de ocorrências BOP
"""

from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional
import os, httpx, json
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

app = FastAPI(title="Agente Inteligente API")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

# ── MODELOS ──────────────────────────────────────────────────

class Message(BaseModel):
    role: str
    content: str

class ChatRequest(BaseModel):
    messages: List[Message]
    system: Optional[str] = "Você é um agente inteligente."
    provider: Optional[str] = "openai"
    model: Optional[str] = None
    max_tokens: Optional[int] = 500

class ChatResponse(BaseModel):
    reply: str
    provider: str
    model: str

class TTSRequest(BaseModel):
    text: str
    voice: Optional[str] = "onyx"
    speed: Optional[float] = 1.0
    model: Optional[str] = "tts-1"

class BOPRequest(BaseModel):
    inserido_por: Optional[str] = ""
    sisgi: Optional[str] = ""
    sgi: Optional[str] = ""
    si_evento: Optional[str] = ""
    operador_ons: Optional[str] = ""
    desligamento: Optional[str] = ""
    nome_equipamento: Optional[str] = ""
    numero_protecao: Optional[str] = ""
    protecao_atuada: Optional[str] = ""
    condicoes_climaticas: Optional[str] = ""
    circuitos_afetados: Optional[str] = ""
    parques_afetados: Optional[str] = ""
    aerogeradores_afetados: Optional[str] = ""
    distancia_rmt: Optional[str] = ""
    data_hora_desligamento: Optional[str] = ""
    data_hora_comunicacao: Optional[str] = ""
    data_hora_energizacao: Optional[str] = ""
    tempo_manobra: Optional[str] = ""
    indisponibilidade: Optional[str] = ""
    causa_detalhada: Optional[str] = ""
    finalizada_por: Optional[str] = ""

# ── GOOGLE SHEETS ─────────────────────────────────────────────

def get_sheets_client():
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        import json as _json

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]

        # Sempre usa o credentials.json na mesma pasta do main.py
        base_dir = os.path.dirname(os.path.abspath(__file__))
        creds_path = os.path.join(base_dir, "credentials.json")

        if not os.path.exists(creds_path):
            raise HTTPException(500, f"credentials.json nao encontrado em: {creds_path}")

        creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
        return gspread.authorize(creds)
    except HTTPException:
        raise
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(500, f"Erro ao conectar Google Sheets: {str(e)}")

EVOLUTION_URL = os.getenv("EVOLUTION_URL", "")
EVOLUTION_KEY = os.getenv("EVOLUTION_KEY", "")
EVOLUTION_INSTANCE = os.getenv("EVOLUTION_INSTANCE", "")
WHATSAPP_GROUP = os.getenv("WHATSAPP_GROUP", "")

# Email config
EMAIL_HOST = os.getenv("EMAIL_HOST", "smtp.gmail.com")
EMAIL_PORT = int(os.getenv("EMAIL_PORT", "587"))
EMAIL_USUARIO = os.getenv("EMAIL_USUARIO", "")
EMAIL_SENHA = os.getenv("EMAIL_SENHA", "")
DESTINATARIOS = os.getenv("EMAIL_DESTINATARIOS", "").split(",")
NOME_REMETENTE = os.getenv("EMAIL_REMETENTE", "Agente BOP")

def enviar_email_sync(row, linha_num):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from datetime import datetime

    def cel(idx): return row[idx] if len(row) > idx and row[idx] else "-"
    def fmt_data(val):
        if not val or val == "-": return "-"
        for fmt in ["%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S"]:
            try: return datetime.strptime(val.strip(), fmt).strftime("%d/%m/%Y %H:%M")
            except: pass
        return val

    equipamento = cel(6)
    desligamento = cel(5)
    protecao = cel(8)
    circuito = cel(10)
    parque = cel(11)
    unidades = cel(12)
    clima = cel(9)
    dt_deslig = fmt_data(cel(14))
    dt_comuni = fmt_data(cel(15))
    dt_energ  = fmt_data(cel(16))
    causa = cel(19)
    inserido = cel(0)
    finalizado = cel(20)
    agora = datetime.now().strftime("%d/%m/%Y %H:%M")

    html = """<!DOCTYPE html>
<html lang="pt-BR">
<head><meta charset="UTF-8"/>
<style>
  body{{margin:0;padding:0;background:#0a0e1a;font-family:Arial,sans-serif}}
  .wrap{{max-width:620px;margin:0 auto;background:#0d1526;border-radius:12px;overflow:hidden;box-shadow:0 8px 32px rgba(0,0,0,.5)}}
  .header{{background:linear-gradient(135deg,#1a2a4a 0%,#0d1a33 100%);padding:32px 36px;border-bottom:3px solid #f59e0b;text-align:center}}
  .header h1{{margin:0;color:#f59e0b;font-size:22px;letter-spacing:3px;text-transform:uppercase}}
  .header p{{margin:8px 0 0;color:#94a3b8;font-size:12px;letter-spacing:2px}}
  .badge{{display:inline-block;background:#ef4444;color:#fff;font-size:11px;font-weight:bold;padding:4px 12px;border-radius:20px;letter-spacing:2px;margin-top:10px}}
  .body{{padding:28px 36px}}
  .section{{margin-bottom:24px}}
  .section-title{{font-size:10px;letter-spacing:3px;color:#f59e0b;text-transform:uppercase;margin-bottom:12px;padding-bottom:6px;border-bottom:1px solid #1e3a5f}}
  .grid{{display:grid;grid-template-columns:1fr 1fr;gap:12px}}
  .field{{background:#111e35;border-radius:8px;padding:12px 14px;border-left:3px solid #1e4d8c}}
  .field.full{{grid-column:1/-1}}
  .field.alert{{border-left-color:#ef4444}}
  .field.ok{{border-left-color:#10b981}}
  .label{{font-size:10px;color:#64748b;letter-spacing:1px;text-transform:uppercase;margin-bottom:4px}}
  .value{{font-size:14px;color:#e2e8f0;font-weight:bold}}
  .footer{{background:#080e1c;padding:18px 36px;text-align:center;border-top:1px solid #1e3a5f}}
  .footer p{{margin:0;color:#334155;font-size:11px}}
  .footer strong{{color:#f59e0b}}
</style>
</head>
<body>
<div class="wrap">
  <div class="header">
    <h1>&#9889; Alerta de Desligamento &#9889;</h1>
    <p>Sistema de Monitoramento Eletrico | Casa dos Ventos</p>
    <span class="badge">NOVO REGISTRO BOP</span>
  </div>
  <div class="body">
    <div class="section">
      <div class="section-title">Equipamento</div>
      <div class="grid">
        <div class="field alert">
          <div class="label">Equipamento</div>
          <div class="value">{equipamento}</div>
        </div>
        <div class="field alert">
          <div class="label">Tipo de Desligamento</div>
          <div class="value">{desligamento}</div>
        </div>
        <div class="field">
          <div class="label">Protecao Atuada</div>
          <div class="value">{protecao}</div>
        </div>
        <div class="field">
          <div class="label">Condicoes Climaticas</div>
          <div class="value">{clima}</div>
        </div>
        <div class="field full">
          <div class="label">Circuito(s) Afetado(s)</div>
          <div class="value">{circuito}</div>
        </div>
        <div class="field">
          <div class="label">Parque Afetado</div>
          <div class="value">{parque}</div>
        </div>
        <div class="field">
          <div class="label">Unidades Geradoras</div>
          <div class="value">{unidades}</div>
        </div>
      </div>
    </div>
    <div class="section">
      <div class="section-title">Cronologia</div>
      <div class="grid">
        <div class="field alert">
          <div class="label">Data/Hora Desligamento</div>
          <div class="value">{dt_deslig}</div>
        </div>
        <div class="field">
          <div class="label">Data/Hora Comunicacao</div>
          <div class="value">{dt_comuni}</div>
        </div>
        <div class="field ok">
          <div class="label">Data/Hora Energizacao</div>
          <div class="value">{dt_energ}</div>
        </div>
      </div>
    </div>
    <div class="section">
      <div class="section-title">Causa</div>
      <div class="grid">
        <div class="field full">
          <div class="label">Causa Detalhada</div>
          <div class="value">{causa}</div>
        </div>
      </div>
    </div>
    <div class="section">
      <div class="section-title">Responsaveis</div>
      <div class="grid">
        <div class="field">
          <div class="label">Inserido por</div>
          <div class="value">{inserido}</div>
        </div>
        <div class="field ok">
          <div class="label">Finalizado por</div>
          <div class="value">{finalizado}</div>
        </div>
      </div>
    </div>
  </div>
  <div class="footer">
    <p>Enviado automaticamente em <strong>{agora}</strong> pelo <strong>Agente Inteligente A-Plan</strong></p>
  </div>
</div>
</body>
</html>""".format(
        equipamento=equipamento, desligamento=desligamento, protecao=protecao,
        circuito=circuito, parque=parque, unidades=unidades, clima=clima,
        dt_deslig=dt_deslig, dt_comuni=dt_comuni, dt_energ=dt_energ,
        causa=causa, inserido=inserido, finalizado=finalizado, agora=agora
    )

    msg = MIMEMultipart("alternative")
    msg["Subject"] = "Alerta BOP - {} - {}".format(equipamento, dt_deslig)
    msg["From"] = "{} <{}>".format(NOME_REMETENTE, EMAIL_USUARIO)
    msg["To"] = ", ".join(DESTINATARIOS)
    msg.attach(MIMEText(html, "html"))

    with smtplib.SMTP(EMAIL_HOST, EMAIL_PORT) as server:
        server.starttls()
        server.login(EMAIL_USUARIO, EMAIL_SENHA)
        server.sendmail(EMAIL_USUARIO, DESTINATARIOS, msg.as_string())
    print("Email enviado para: {}".format(DESTINATARIOS))

async def enviar_whatsapp_bop(ws, linha_num):
    import time
    time.sleep(2)
    row = ws.row_values(linha_num)
    def cel(idx): return row[idx] if len(row) > idx and row[idx] else "-"
    def fmt_data(val):
        if not val or val == "-": return "-"
        from datetime import datetime
        for fmt in ["%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S"]:
            try: return datetime.strptime(val.strip(), fmt).strftime("%d/%m/%Y %H:%M")
            except: pass
        return val
    sep = "―" * 22
    linhas = [
        "⚡ *ALERTA DE DESLIGAMENTO* ⚡",
        sep, "",
        "🔧 *Equipamento:* {}".format(cel(6)),
        "💥 *Desligamento:* {}".format(cel(5)),
        "🚨 *Proteção:* {}".format(cel(8)),
        "📍 *Circuito:* {}".format(cel(10)),
        "🌱 *Parque Afetado:* {}".format(cel(11)),
        "💨 *Unidades Geradoras:* {}".format(cel(12)),
        "☔ *Clima:* {}".format(cel(9)),
        "🕒 *Data/Hora:* {}".format(fmt_data(cel(14))),
        "📝 *Causa:* {}".format(cel(19)),
        "", sep,
        "👤 *Inserido por:* {}".format(cel(0)),
        "✅ *Finalizado por:* {}".format(cel(20)),
    ]
    mensagem = "\n".join(linhas)
    async with httpx.AsyncClient(timeout=10) as c:
        r = await c.post(
            "{}/message/sendText/{}".format(EVOLUTION_URL, EVOLUTION_INSTANCE),
            headers={"apikey": EVOLUTION_KEY, "Content-Type": "application/json"},
            json={"number": WHATSAPP_GROUP, "text": mensagem}
        )
        print("WhatsApp status: {}".format(r.status_code))

@app.post("/bop")
async def inserir_bop(req: BOPRequest):
    try:
        gc = get_sheets_client()
        # tenta variavel de ambiente, senao usa o ID fixo
        sheet_id = os.getenv("SHEET_ID", "") or "1K2EKqHPErZEr-MiT8CWOhDazMlq0gbsyfzZDp1AAUAc"
        if not sheet_id:
            raise HTTPException(400, "SHEET_ID nao configurado")

        sh = gc.open_by_key(sheet_id)
        ws = sh.worksheet("BOP")

        linha = [
            req.inserido_por,
            req.sisgi,
            req.sgi,
            req.si_evento,
            req.operador_ons,
            req.desligamento,
            req.nome_equipamento,
            req.numero_protecao,
            req.protecao_atuada,
            req.condicoes_climaticas,
            req.circuitos_afetados,
            req.parques_afetados,
            req.aerogeradores_afetados,
            req.distancia_rmt,
            req.data_hora_desligamento,
            req.data_hora_comunicacao,
            req.data_hora_energizacao,
            req.tempo_manobra,
            req.indisponibilidade,
            req.causa_detalhada,
            req.finalizada_por,
        ]

        # formata datas para dd/mm/yyyy hh:mm
        from datetime import datetime
        def formatar_data(val):
            if not val: return val
            for fmt in ["%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"]:
                try: return datetime.strptime(val.strip(), fmt).strftime("%d/%m/%Y %H:%M")
                except: pass
            return val

        linha[14] = formatar_data(linha[14])  # data_hora_desligamento
        linha[15] = formatar_data(linha[15])  # data_hora_comunicacao
        linha[16] = formatar_data(linha[16])  # data_hora_energizacao

        # usa coluna O (Data/Hora Desligamento) para achar primeira linha sem registro real
        col_o = ws.col_values(15)  # coluna O
        primeira_vazia = 2  # começa na linha 2 (pula cabeçalho)
        for i, val in enumerate(col_o[1:], start=2):
            if not val.strip():
                primeira_vazia = i
                break
        else:
            primeira_vazia = len(col_o) + 1

        # coluna H = N° protecao:
        # se tiver numero informado usa, senao coloca "N"
        num_protecao = linha[7] if linha[7] and linha[7].strip() else "N"

        # monta cada celula individualmente — preserva formulas em K,L,M,N,R,S
        cells = [
            {'range': f'B{primeira_vazia}', 'values': [[linha[1]]]},   # N SI
            {'range': f'C{primeira_vazia}', 'values': [[linha[2]]]},   # SGI
            {'range': f'D{primeira_vazia}', 'values': [[linha[3]]]},   # N Evento
            {'range': f'E{primeira_vazia}', 'values': [[linha[4]]]},   # Operador ONS
            {'range': f'F{primeira_vazia}', 'values': [[linha[5]]]},   # Desligamento
            {'range': f'G{primeira_vazia}', 'values': [[linha[6]]]},   # Nome Equipamento
            {'range': f'H{primeira_vazia}', 'values': [[num_protecao]]}, # N Protecao
            # I = formula VLOOKUP — nao escreve
            {'range': f'J{primeira_vazia}', 'values': [[linha[9]]]},   # Condicoes Climaticas
            # K,L,M,N = formulas — nao escreve
            {'range': f'O{primeira_vazia}', 'values': [[linha[14]]]},  # Data Desligamento
            {'range': f'P{primeira_vazia}', 'values': [[linha[15]]]},  # Data Comunicacao
            {'range': f'Q{primeira_vazia}', 'values': [[linha[16]]]},  # Data Energizacao
            # R,S = formulas — nao escreve
            {'range': f'T{primeira_vazia}', 'values': [[linha[19]]]},  # Causa Detalhada
            {'range': f'U{primeira_vazia}', 'values': [[linha[20]]]},  # Finalizada Por
        ]
        ws.batch_update(cells, value_input_option='USER_ENTERED')

        # coluna A — usa update_acell que funciona mesmo com lista suspensa
        print(f"DEBUG coluna A linha[0]: '{linha[0]}' na linha {primeira_vazia}")
        ws.update_acell(f'A{primeira_vazia}', linha[0])
        print(f"DEBUG coluna A atualizada!")

        # envia notificacao WhatsApp
        wpp_status = "ok"
        email_status = "ok"
        try:
            await enviar_whatsapp_bop(ws, primeira_vazia)
        except Exception as e:
            wpp_status = "erro: {}".format(str(e)[:60])
            print("WhatsApp erro: {}".format(e))
        try:
            import asyncio, concurrent.futures
            row_data = ws.row_values(primeira_vazia)
            loop = asyncio.get_event_loop()
            await loop.run_in_executor(None, enviar_email_sync, row_data, primeira_vazia)
        except Exception as e:
            email_status = "erro: {}".format(str(e)[:60])
            print("Email erro: {}".format(e))

        mensagem = (
            "Registro inserido na planilha (linha {}).\n"
            "WhatsApp: {}.\n"
            "Email: {}."
        ).format(primeira_vazia, wpp_status, email_status)

        return {"status": "ok", "mensagem": mensagem}

    except HTTPException:
        raise
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(500, f"Erro ao inserir na planilha: {str(e)}")

@app.get("/bop/status")
async def bop_status():
    try:
        gc = get_sheets_client()
        sheet_id = os.getenv("SHEET_ID", "")
        if not sheet_id:
            return {"status": "erro", "mensagem": "SHEET_ID não configurado"}
        sh = gc.open_by_key(sheet_id)
        ws = sh.worksheet("BOP")
        linhas = len(ws.get_all_values())
        return {"status": "ok", "linhas": linhas, "mensagem": f"Planilha BOP acessível — {linhas} linhas"}
    except Exception as e:
        return {"status": "erro", "mensagem": str(e)}

# ── CHAT ─────────────────────────────────────────────────────

async def call_openai(messages, system, model, max_tokens):
    key = os.getenv("OPENAI_API_KEY")
    if not key: raise HTTPException(400, "OPENAI_API_KEY não configurada no .env")
    model = model or "gpt-4o-mini"
    async with httpx.AsyncClient(timeout=60) as c:
        r = await c.post("https://api.openai.com/v1/chat/completions",
            headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
            json={"model": model, "max_tokens": max_tokens,
                  "messages": [{"role":"system","content":system}]+[m.dict() for m in messages]})
        if r.status_code != 200: raise HTTPException(r.status_code, f"OpenAI: {r.text}")
        return r.json()["choices"][0]["message"]["content"], model

async def call_claude(messages, system, model, max_tokens):
    key = os.getenv("ANTHROPIC_API_KEY")
    if not key: raise HTTPException(400, "ANTHROPIC_API_KEY não configurada no .env")
    model = model or "claude-haiku-4-5-20251001"
    async with httpx.AsyncClient(timeout=60) as c:
        r = await c.post("https://api.anthropic.com/v1/messages",
            headers={"x-api-key": key, "anthropic-version": "2023-06-01", "Content-Type": "application/json"},
            json={"model": model, "max_tokens": max_tokens, "system": system,
                  "messages": [m.dict() for m in messages]})
        if r.status_code != 200: raise HTTPException(r.status_code, f"Claude: {r.text}")
        return r.json()["content"][0]["text"], model

async def call_ollama(messages, system, model, max_tokens):
    url = os.getenv("OLLAMA_URL", "http://localhost:11434")
    model = model or "llama3.2"
    async with httpx.AsyncClient(timeout=120) as c:
        try:
            r = await c.post(f"{url}/api/chat",
                json={"model": model, "stream": False,
                      "messages": [{"role":"system","content":system}]+[m.dict() for m in messages]})
        except httpx.ConnectError:
            raise HTTPException(503, f"Ollama offline em {url}")
        if r.status_code != 200: raise HTTPException(r.status_code, f"Ollama: {r.text}")
        return r.json()["message"]["content"], model

@app.post("/chat", response_model=ChatResponse)
async def chat(req: ChatRequest):
    from datetime import datetime
    data_atual = datetime.now().strftime("%d/%m/%Y")
    hora_atual = datetime.now().strftime("%H:%M")
    # injeta data/hora atual no system prompt
    system_com_data = f"Data de hoje: {data_atual}. Hora atual: {hora_atual}.\n\n{req.system}"
    p = req.provider.lower()
    if p == "openai":   reply, model = await call_openai(req.messages, system_com_data, req.model, req.max_tokens)
    elif p == "claude": reply, model = await call_claude(req.messages, system_com_data, req.model, req.max_tokens)
    elif p == "ollama": reply, model = await call_ollama(req.messages, system_com_data, req.model, req.max_tokens)
    else: raise HTTPException(400, f"Provider inválido: {p}")
    return ChatResponse(reply=reply, provider=p, model=model)

# ── TTS ──────────────────────────────────────────────────────

@app.post("/tts")
async def tts(req: TTSRequest):
    key = os.getenv("OPENAI_API_KEY")
    if not key: raise HTTPException(400, "OPENAI_API_KEY não configurada no .env")
    voices = ["alloy","echo","fable","onyx","nova","shimmer"]
    voice = req.voice if req.voice in voices else "onyx"
    async with httpx.AsyncClient(timeout=30) as c:
        r = await c.post("https://api.openai.com/v1/audio/speech",
            headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
            json={"model": req.model or "tts-1", "input": req.text[:4096],
                  "voice": voice, "speed": max(0.25, min(4.0, req.speed or 1.0))})
        if r.status_code != 200: raise HTTPException(r.status_code, f"TTS: {r.text}")
        return StreamingResponse(iter([r.content]), media_type="audio/mpeg")

# ── STATUS ───────────────────────────────────────────────────

@app.get("/status")
async def status():
    p = {}
    p["openai"] = "configurado" if os.getenv("OPENAI_API_KEY") else "sem chave"
    p["claude"] = "configurado" if os.getenv("ANTHROPIC_API_KEY") else "sem chave"
    try:
        async with httpx.AsyncClient(timeout=3) as c:
            r = await c.get(os.getenv("OLLAMA_URL","http://localhost:11434"))
            p["ollama"] = "online" if r.status_code==200 else "offline"
    except: p["ollama"] = "offline"
    # testa sheets
    creds_ok = os.path.exists(os.getenv("GOOGLE_CREDENTIALS","credentials.json"))
    sheet_ok  = bool(os.getenv("SHEET_ID",""))
    p["sheets"] = "configurado" if (creds_ok and sheet_ok) else "não configurado"
    return {"status": "ok", "providers": p}

@app.get("/ollama/models")
async def ollama_models():
    url = os.getenv("OLLAMA_URL","http://localhost:11434")
    try:
        async with httpx.AsyncClient(timeout=5) as c:
            r = await c.get(f"{url}/api/tags")
            return {"models": [m["name"] for m in r.json().get("models",[])]}
    except: return {"models": []}

# ── STATIC ───────────────────────────────────────────────────

app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/")
async def root():
    return FileResponse("static/index.html")

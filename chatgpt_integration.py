import os
import json
import urllib.request
import urllib.error

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY") or os.getenv("CHATGPT_API_KEY")
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")

CHATGPT_HABILITADO = bool(OPENAI_API_KEY)


def validar_chave_api():
    return bool(OPENAI_API_KEY)


def _base_para_texto(base):
    if base is None:
        return ""

    if isinstance(base, str):
        return base

    if isinstance(base, list):
        partes = []
        for item in base:
            if isinstance(item, dict):
                pergunta = item.get("question") or item.get("q") or item.get("titulo") or ""
                resposta = item.get("answer") or item.get("a") or item.get("conteudo") or ""
                partes.append(f"Pergunta/Título: {pergunta}\nResposta/Conteúdo: {resposta}")
            else:
                partes.append(str(item))
        return "\n\n".join(partes)

    if isinstance(base, dict):
        partes = []
        for chave, valor in base.items():
            partes.append(f"{chave}: {valor}")
        return "\n\n".join(partes)

    return str(base)


def analisar_com_chatgpt(pergunta, contexto=None):
    if not OPENAI_API_KEY:
        return None

    contexto_texto = _base_para_texto(contexto)

    system_prompt = (
        "Você é a ZIA, assistente virtual da Opus.\n"
        "Responda em português do Brasil.\n"
        "Responda com base APENAS nas informações fornecidas no contexto.\n"
        "Se a informação não estiver no contexto, diga que não encontrou informação suficiente na base cadastrada.\n"
        "Seja direta, objetiva e profissional.\n\n"
        "Contexto disponível:\n"
        f"{contexto_texto}"
    )

    payload = {
        "model": OPENAI_MODEL,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": pergunta}
        ],
        "temperature": 0.2,
        "max_tokens": 700
    }

    req = urllib.request.Request(
        "https://api.openai.com/v1/chat/completions",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {OPENAI_API_KEY}"
        },
        method="POST"
    )

    try:
        with urllib.request.urlopen(req, timeout=45) as response:
            raw = response.read().decode("utf-8")
            data = json.loads(raw)

        choices = data.get("choices", [])
        if not choices:
            return None

        return choices[0]["message"]["content"].strip()

    except urllib.error.HTTPError as e:
        try:
            erro = e.read().decode("utf-8")
        except Exception:
            erro = str(e)
        print(f"[ERRO OPENAI HTTP] {e.code}: {erro}")
        return None

    except Exception as e:
        print(f"[ERRO OPENAI] {type(e).__name__}: {e}")
        return None


def gerar_resposta_com_fallback(pergunta, base=None, fallback_fn=None):
    contexto_texto = _base_para_texto(base)

    if CHATGPT_HABILITADO:
        resposta = analisar_com_chatgpt(pergunta, contexto_texto)
        if resposta:
            return resposta

    if fallback_fn:
        try:
            return fallback_fn(pergunta, base)
        except Exception as e:
            print(f"[ERRO FALLBACK] {e}")

    if contexto_texto:
        return contexto_texto[:1500]

    return None


def interpretar_pergunta(pergunta):
    return {
        "pergunta": pergunta,
        "termos": str(pergunta).lower().split()
    }

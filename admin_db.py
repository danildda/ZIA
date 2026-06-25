import os
import json
import threading
import urllib.request
import urllib.parse

# admin_db compatibility layer
# If CF_API_URL is set in environment, this module will forward requests to that endpoint.
# Otherwise it uses a simple JSON file store in the same directory as this file.

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STORE_PATH = os.path.join(BASE_DIR, 'admin_db_store.json')
CF_API_URL = os.getenv('CF_API_URL')  # e.g. https://meu-worker.example.workers.dev
LOCK = threading.Lock()


def _init_default_store():
    return {
        'next_id': 1,
        'modules': [
            {'id': 1, 'nome': 'qualidade'},
            {'id': 2, 'nome': 'orcamento'},
            {'id': 3, 'nome': 'texto'}
        ],
        'conteudos': {}
    }


def _load_store():
    if CF_API_URL:
        return None
    if not os.path.exists(STORE_PATH):
        store = _init_default_store()
        _save_store(store)
        return store
    try:
        with open(STORE_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        store = _init_default_store()
        _save_store(store)
        return store


def _save_store(store):
    if CF_API_URL:
        return
    with LOCK:
        with open(STORE_PATH, 'w', encoding='utf-8') as f:
            json.dump(store, f, ensure_ascii=False, indent=2)


def _cf_request(path, method='GET', data=None):
    if not CF_API_URL:
        raise RuntimeError('CF_API_URL not configured')
    url = CF_API_URL.rstrip('/') + '/' + path.lstrip('/')
    headers = {'Content-Type': 'application/json'}
    body = None
    if data is not None:
        body = json.dumps(data).encode('utf-8')
    req = urllib.request.Request(url, data=body, headers=headers, method=method)
    with urllib.request.urlopen(req, timeout=30) as resp:
        resp_body = resp.read()
        try:
            return json.loads(resp_body.decode('utf-8'))
        except Exception:
            return {}


def init_db():
    """Initialize database (no-op for Cloudflare; creates local store otherwise)."""
    if CF_API_URL:
        # Optionally, could ping the endpoint to ensure availability
        return
    _load_store()


def get_modulos():
    if CF_API_URL:
        return _cf_request('/modulos') or []
    store = _load_store()
    return store.get('modules', [])


def get_conteudos_modulo(modulo_id):
    # Accept either int id or module name
    if CF_API_URL:
        return _cf_request(f'/conteudos?modulo_id={urllib.parse.quote(str(modulo_id))}') or []
    store = _load_store()
    # Resolve name if needed
    modulo_name = None
    if isinstance(modulo_id, int):
        for m in store.get('modules', []):
            if m.get('id') == modulo_id:
                modulo_name = m.get('nome')
                break
    else:
        modulo_name = str(modulo_id)
    resultado = []
    for cid, item in store.get('conteudos', {}).items():
        if item.get('modulo') == modulo_name or item.get('modulo_id') == modulo_id:
            resultado.append({'id': int(cid), 'nome': item.get('nome', ''), 'descricao': item.get('descricao', '')})
    return resultado


def get_dados_conteudo(conteudo_id):
    if CF_API_URL:
        return _cf_request(f'/dados/{conteudo_id}') or {}
    store = _load_store()
    return store.get('conteudos', {}).get(str(conteudo_id), {})


def criar_conteudo(modulo_id, nome, descricao, tag):
    if CF_API_URL:
        return _cf_request('/criar', method='POST', data={'modulo_id': modulo_id, 'nome': nome, 'descricao': descricao, 'tag': tag}) or (False, 'Erro')
    store = _load_store()
    with LOCK:
        nid = store.get('next_id', 1)
        store['next_id'] = nid + 1
        # resolve module name
        modulo_name = None
        if isinstance(modulo_id, int):
            for m in store.get('modules', []):
                if m.get('id') == modulo_id:
                    modulo_name = m.get('nome')
                    break
        else:
            modulo_name = str(modulo_id)
        store['conteudos'][str(nid)] = {
            'id': nid,
            'modulo': modulo_name,
            'modulo_id': modulo_id,
            'nome': nome,
            'descricao': descricao,
            'tag': tag,
            'texto': '',
            'arquivo': ''
        }
        _save_store(store)
    return True, 'Conteúdo criado com sucesso'


def atualizar_dados_conteudo(conteudo_id, texto):
    if CF_API_URL:
        return _cf_request(f'/atualizar/{conteudo_id}', method='POST', data={'texto': texto}) or (False, 'Erro')
    store = _load_store()
    cid = str(conteudo_id)
    if cid not in store.get('conteudos', {}):
        return False, 'Conteúdo não encontrado'
    store['conteudos'][cid]['texto'] = texto
    _save_store(store)
    return True, 'Texto atualizado'


def atualizar_tag_conteudo(conteudo_id, tag):
    if CF_API_URL:
        return _cf_request(f'/atualizar/{conteudo_id}', method='POST', data={'tag': tag}) or (False, 'Erro')
    store = _load_store()
    cid = str(conteudo_id)
    if cid not in store.get('conteudos', {}):
        return False, 'Conteúdo não encontrado'
    store['conteudos'][cid]['tag'] = tag
    _save_store(store)
    return True, 'Tag atualizada'


def atualizar_arquivo_conteudo(conteudo_id, filename):
    if CF_API_URL:
        return _cf_request(f'/atualizar/{conteudo_id}', method='POST', data={'arquivo': filename}) or (False, 'Erro')
    store = _load_store()
    cid = str(conteudo_id)
    if cid not in store.get('conteudos', {}):
        return False, 'Conteúdo não encontrado'
    store['conteudos'][cid]['arquivo'] = filename
    _save_store(store)
    return True, 'Arquivo atualizado'


def deletar_conteudo(conteudo_id):
    if CF_API_URL:
        return _cf_request(f'/deletar/{conteudo_id}', method='POST') or (False, 'Erro')
    store = _load_store()
    cid = str(conteudo_id)
    if cid in store.get('conteudos', {}):
        del store['conteudos'][cid]
        _save_store(store)
        return True, 'Conteúdo deletado'
    return False, 'Conteúdo não encontrado'


def buscar_conteudo_modulo(modulo, key):
    """
    Retorna um dicionário com possíveis chaves: 'sucesso', 'texto', 'arquivo', 'nome'
    If multiple matches, prefer exact nome/tag match; otherwise return first partial match.
    """
    if CF_API_URL:
        return _cf_request(f'/buscar?modulo={urllib.parse.quote(str(modulo))}&key={urllib.parse.quote(str(key))}') or {}
    store = _load_store()
    modulo_name = str(modulo)
    # search conteudos in module
    matches = []
    for item in store.get('conteudos', {}).values():
        if item.get('modulo') != modulo_name:
            continue
        # exact match on nome or tag
        if str(item.get('nome','')).lower() == str(key).lower() or str(item.get('tag','')).lower() == str(key).lower():
            return {'sucesso': True, 'texto': item.get('texto', ''), 'arquivo': item.get('arquivo', ''), 'nome': item.get('nome','')}
        # partial match
        if str(key).lower() in str(item.get('nome','')).lower() or str(key).lower() in str(item.get('texto','')).lower() or str(key).lower() in str(item.get('tag','')).lower():
            matches.append(item)
    if matches:
        # combine texts
        combined = '\n\n'.join([m.get('texto','') for m in matches if m.get('texto')])
        return {'sucesso': True, 'texto': combined, 'arquivo': matches[0].get('arquivo',''), 'nome': matches[0].get('nome','')}
    return {'sucesso': False}


def buscar_conteudos_modulo_por_query(modulo, query):
    if CF_API_URL:
        return _cf_request(f'/buscar-multiplo?modulo={urllib.parse.quote(str(modulo))}&q={urllib.parse.quote(str(query))}') or []
    store = _load_store()
    modulo_name = str(modulo)
    results = []
    for item in store.get('conteudos', {}).values():
        if item.get('modulo') != modulo_name:
            continue
        if query.lower() in (item.get('nome','') or '').lower() or query.lower() in (item.get('texto','') or '').lower() or query.lower() in (item.get('tag','') or '').lower():
            results.append((item.get('nome',''), item.get('texto','')))
    return results

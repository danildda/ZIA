// Exemplo de Cloudflare Worker para armazenar os dados do admin
// Requer um KV namespace vinculado como 'ADMIN_DB'
addEventListener('fetch', event => {
  event.respondWith(handleRequest(event.request))
})

async function getStore() {
  const raw = await ADMIN_DB.get('store')
  if (!raw) {
    const init = {
      next_id: 1,
      modules: [
        {id: 1, nome: 'qualidade'},
        {id: 2, nome: 'orcamento'},
        {id: 3, nome: 'texto'}
      ],
      conteudos: {}
    }
    await ADMIN_DB.put('store', JSON.stringify(init))
    return init
  }
  return JSON.parse(raw)
}

async function saveStore(store) {
  await ADMIN_DB.put('store', JSON.stringify(store))
}

async function handleRequest(request) {
  const url = new URL(request.url)
  const path = url.pathname
  try {
    if (path === '/modulos' && request.method === 'GET') {
      const s = await getStore()
      return new Response(JSON.stringify(s.modules), {headers: {'Content-Type':'application/json'}})
    }

    if (path === '/conteudos' && request.method === 'GET') {
      const s = await getStore()
      const modulo_q = url.searchParams.get('modulo_id') || url.searchParams.get('modulo')
      const results = []
      for (const [k,v] of Object.entries(s.conteudos)){
        if (!modulo_q) { results.push(v); continue }
        if (String(v.modulo) === modulo_q || String(v.modulo_id) === modulo_q) results.push(v)
      }
      return new Response(JSON.stringify(results), {headers: {'Content-Type':'application/json'}})
    }

    if (path.startsWith('/dados/') && request.method === 'GET') {
      const id = path.split('/')[2]
      const s = await getStore()
      return new Response(JSON.stringify(s.conteudos[id] || {}), {headers: {'Content-Type':'application/json'}})
    }

    if (path === '/criar' && request.method === 'POST') {
      const payload = await request.json()
      const s = await getStore()
      const nid = s.next_id++
      s.conteudos[String(nid)] = {
        id: nid,
        modulo: payload.modulo || payload.modulo_id || payload.modulo_nome || '',
        modulo_id: payload.modulo_id || payload.modulo || '',
        nome: payload.nome || '',
        descricao: payload.descricao || '',
        tag: payload.tag || '',
        texto: '',
        arquivo: ''
      }
      await saveStore(s)
      return new Response(JSON.stringify([true, 'Criado']), {headers: {'Content-Type':'application/json'}})
    }

    if (path.startsWith('/atualizar/') && request.method === 'POST') {
      const id = path.split('/')[2]
      const payload = await request.json()
      const s = await getStore()
      if (!s.conteudos[String(id)]) return new Response(JSON.stringify([false,'Não encontrado']), {headers:{'Content-Type':'application/json'}})
      Object.assign(s.conteudos[String(id)], payload)
      await saveStore(s)
      return new Response(JSON.stringify([true,'Atualizado']), {headers:{'Content-Type':'application/json'}})
    }

    if (path.startsWith('/deletar/') && request.method === 'POST') {
      const id = path.split('/')[2]
      const s = await getStore()
      if (s.conteudos[String(id)]) delete s.conteudos[String(id)]
      await saveStore(s)
      return new Response(JSON.stringify([true,'Deletado']), {headers:{'Content-Type':'application/json'}})
    }

    if (path === '/buscar' && request.method === 'GET') {
      const modulo = url.searchParams.get('modulo') || ''
      const key = url.searchParams.get('key') || ''
      const s = await getStore()
      const items = Object.values(s.conteudos).filter(i => i.modulo === modulo)
      for (const it of items) {
        if ((it.nome||'').toLowerCase() === key.toLowerCase() || (it.tag||'').toLowerCase() === key.toLowerCase()) {
          return new Response(JSON.stringify({sucesso:true, texto: it.texto||'', arquivo: it.arquivo||'', nome: it.nome||''}), {headers:{'Content-Type':'application/json'}})
        }
      }
      // partial matches
      const matches = items.filter(it => (it.nome||'').toLowerCase().includes(key.toLowerCase()) || (it.texto||'').toLowerCase().includes(key.toLowerCase()) || (it.tag||'').toLowerCase().includes(key.toLowerCase()))
      if (matches.length) {
        const combined = matches.map(m=>m.texto||'').join('\n\n')
        return new Response(JSON.stringify({sucesso:true, texto: combined, arquivo: matches[0].arquivo||'', nome: matches[0].nome||''}), {headers:{'Content-Type':'application/json'}})
      }
      return new Response(JSON.stringify({sucesso:false}), {headers:{'Content-Type':'application/json'}})
    }

    if (path === '/buscar-multiplo' && request.method === 'GET') {
      const modulo = url.searchParams.get('modulo') || ''
      const q = url.searchParams.get('q') || ''
      const s = await getStore()
      const items = Object.values(s.conteudos).filter(i => i.modulo === modulo)
      const results = items.filter(it => (it.nome||'').toLowerCase().includes(q.toLowerCase()) || (it.texto||'').toLowerCase().includes(q.toLowerCase()) || (it.tag||'').toLowerCase().includes(q.toLowerCase())).map(it=>[it.nome, it.texto])
      return new Response(JSON.stringify(results), {headers:{'Content-Type':'application/json'}})
    }

    return new Response('Not Found', {status:404})
  } catch (e) {
    return new Response(JSON.stringify({error: String(e)}), {status:500, headers:{'Content-Type':'application/json'}})
  }
}

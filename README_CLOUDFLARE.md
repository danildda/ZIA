Cloudflare Worker for ZIA admin (example)

Overview
--------
This folder contains an example Cloudflare Worker (`worker_admin_db.js`) that implements a simple REST API
using a KV namespace called `ADMIN_DB`. The Python app `ZIA WEB R00.py` includes a local shim `admin_db.py`
that will forward requests to the Worker when the environment variable `CF_API_URL` is set.

Quick steps to deploy
---------------------
1. Install `wrangler` (Cloudflare CLI) and log in:

```bash
# install wrangler if you don't have it
npm install -g wrangler
wrangler login
```

2. Create a new Worker project and copy `worker_admin_db.js` into it, or simply configure `wrangler.toml` to point to this file.

3. Create a KV namespace in Cloudflare Dashboard and add it to `wrangler.toml`:

```toml
name = "zia-admin-db"

[[kv_namespaces]]
binding = "ADMIN_DB"
id = "<your-kv-id>"
```

4. Publish the Worker:

```bash
wrangler publish
```

5. Set `CF_API_URL` in the environment where you run the Python app to the Worker URL, e.g. `https://my-worker.example.workers.dev`.

Notes
-----
- The Worker uses a single KV key `store` to keep a JSON object with modules and contents. This is a minimal example;
  for production consider using Durable Objects or D1 for better query capabilities and concurrency control.
- The Python shim will fallback to a local JSON file `admin_db_store.json` if `CF_API_URL` is not set, so you can test locally without Cloudflare.

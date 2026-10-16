#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dashboard de remoção (foto principal + galeria + downloads) por CSV.

Precisa estar NA MESMA PASTA do remover_tudo_woocommerce.py
(usa as funções dele, inclusive o modo rápido que pula produto vazio).

Instalar (uma vez):  py -m pip install flask playwright
Rodar:               py remover_dashboard.py
Abre sozinho em:     http://localhost:8788
"""

import io
import threading
import webbrowser
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from playwright.sync_api import sync_playwright

import remover_tudo_woocommerce as rw

SCRIPT_DIR = Path(__file__).resolve().parent
PORTA = 8788

SITES = {
    "sag": ("https://sagloja.devsplcomunic.com.br", "chrome-profile-sag"),
    "choice": ("https://choicetech.com.br", "chrome-profile-choice"),
}

app = Flask(__name__)
lock = threading.Lock()
estado = {
    "rodando": False, "parar": False, "logs": [], "total": 0, "feitos": 0,
    "ok": 0, "pulado": 0, "erro": 0, "atual": "", "resultados": [],
}


def dlog(msg: str):
    hora = datetime.now().strftime("%H:%M:%S")
    with lock:
        estado["logs"].append(f"[{hora}] {msg}")
    print(msg, flush=True)


rw.log = dlog  # todos os logs do script vão pro painel


def worker(skus: list[str], site: str):
    base, perfil = SITES[site]
    rw.BASE_URL = base
    rw.PROFILE_DIR = SCRIPT_DIR / perfil

    try:
        with sync_playwright() as p:
            context = p.chromium.launch_persistent_context(
                user_data_dir=str(rw.PROFILE_DIR),
                headless=False,
                viewport={"width": 1440, "height": 900},
            )
            page = context.pages[0] if context.pages else context.new_page()
            dlog(f"Site: {base}")
            rw.wait_for_login(page)

            for i, sku in enumerate(skus, start=1):
                if estado["parar"]:
                    dlog("PARADO pelo usuário.")
                    break
                with lock:
                    estado["atual"] = sku
                dlog(f"[{i}/{len(skus)}] SKU {sku}")
                try:
                    status = rw.process_sku(page, sku)
                    detalhe = ""
                except Exception as e:
                    status, detalhe = "ERRO", str(e)
                    dlog(f"  !! ERRO: {detalhe}")
                    try:
                        dbg = SCRIPT_DIR / "debug_remocao_fotos"
                        dbg.mkdir(exist_ok=True)
                        page.screenshot(path=str(dbg / f"erro_{sku}.png"), full_page=True)
                    except Exception:
                        pass
                with lock:
                    estado["resultados"].append((sku, status, detalhe))
                    estado["feitos"] += 1
                    chave = {"OK": "ok", "PULADO": "pulado"}.get(status, "erro")
                    estado[chave] += 1

            context.close()
    except Exception as e:
        dlog(f"!! Falha geral: {e}")
    finally:
        with lock:
            estado["rodando"] = False
            estado["atual"] = ""
        dlog(f"=== FINALIZADO === OK: {estado['ok']} | PULADOS: {estado['pulado']} | ERROS: {estado['erro']}")


@app.post("/iniciar")
def iniciar():
    if estado["rodando"]:
        return jsonify(erro="Já está rodando."), 400

    site = request.form.get("site", "sag")
    if site not in SITES:
        return jsonify(erro="Site inválido."), 400

    tmp = SCRIPT_DIR / "_lista_dashboard.csv"
    arq = request.files.get("csv")
    texto = (request.form.get("texto") or "").strip()
    if arq and arq.filename:
        arq.save(tmp)
    elif texto:
        tmp.write_text(texto, encoding="utf-8")
    else:
        return jsonify(erro="Envie um CSV ou cole os SKUs."), 400

    skus = rw.carrega_skus(tmp)
    if not skus:
        return jsonify(erro="Nenhum SKU encontrado."), 400

    with lock:
        estado.update(rodando=True, parar=False, logs=[], total=len(skus), feitos=0,
                      ok=0, pulado=0, erro=0, atual="", resultados=[])
    dlog(f"{len(skus)} SKU(s) carregados.")
    threading.Thread(target=worker, args=(skus, site), daemon=True).start()
    return jsonify(ok=True, total=len(skus))


@app.post("/parar")
def parar():
    estado["parar"] = True
    dlog("Parando após o SKU atual...")
    return jsonify(ok=True)


@app.get("/status")
def status():
    desde = int(request.args.get("desde", 0))
    with lock:
        return jsonify(
            rodando=estado["rodando"], total=estado["total"], feitos=estado["feitos"],
            ok=estado["ok"], pulado=estado["pulado"], erro=estado["erro"],
            atual=estado["atual"], logs=estado["logs"][desde:], n=len(estado["logs"]),
            tem_relatorio=bool(estado["resultados"]),
        )


@app.get("/relatorio")
def relatorio():
    buf = io.StringIO()
    buf.write("sku;status;detalhe\n")
    for sku, st, det in estado["resultados"]:
        buf.write(f"{sku};{st};{det.replace(';', ',')}\n")
    data = io.BytesIO(buf.getvalue().encode("utf-8-sig"))
    nome = f"relatorio_remocao_{datetime.now():%Y%m%d_%H%M}.csv"
    return send_file(data, mimetype="text/csv", as_attachment=True, download_name=nome)


@app.get("/")
def index():
    return HTML


HTML = """<!doctype html><html lang="pt-br"><head><meta charset="utf-8">
<title>Remover fotos/downloads - WooCommerce</title>
<style>
*{box-sizing:border-box}body{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f3f4f6;color:#111}
.w{max-width:960px;margin:24px auto;padding:0 16px}h1{font-size:20px;margin:0 0 16px}
.c{background:#fff;border-radius:10px;padding:16px;margin-bottom:14px;box-shadow:0 1px 3px #0001}
#drop{border:2px dashed #9ca3af;border-radius:8px;padding:22px;text-align:center;cursor:pointer;color:#555}
#drop.on{border-color:#2563eb;background:#eff6ff}
textarea{width:100%;height:90px;margin-top:10px;font-family:Consolas,monospace;font-size:13px;padding:8px}
.row{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-top:12px}
select,button{padding:8px 14px;border-radius:6px;border:1px solid #d1d5db;font-size:14px}
button{cursor:pointer;background:#2563eb;color:#fff;border:0}button:disabled{opacity:.5;cursor:default}
#parar{background:#dc2626}#rel{background:#059669}
.k{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}
.k div{background:#fff;border-radius:10px;padding:12px;text-align:center;box-shadow:0 1px 3px #0001}
.k b{display:block;font-size:24px}.ok b{color:#059669}.pu b{color:#6b7280}.er b{color:#dc2626}
.bar{height:8px;background:#e5e7eb;border-radius:4px;overflow:hidden;margin-top:12px}
.bar i{display:block;height:100%;width:0;background:#2563eb;transition:width .3s}
#log{background:#111827;color:#d1d5db;font-family:Consolas,monospace;font-size:12.5px;height:380px;overflow:auto;padding:10px;border-radius:8px;white-space:pre-wrap}
</style></head><body><div class="w">
<h1>Remover foto principal / galeria / downloads (só SKUs do CSV)</h1>
<div class="c">
  <div id="drop">Arraste o CSV aqui ou clique para escolher<br><small id="fn"></small></div>
  <input type="file" id="f" accept=".csv,.txt" hidden>
  <textarea id="t" placeholder="...ou cole os SKUs aqui, um por linha"></textarea>
  <div class="row">
    <label>Site: <select id="site"><option value="sag">SAG (sagloja.devsplcomunic)</option><option value="choice">Choice (choicetech.com.br)</option></select></label>
    <button id="ini">Iniciar remoção</button>
    <button id="parar" disabled>Parar</button>
    <button id="rel" disabled>Baixar relatório</button>
  </div>
  <div class="bar"><i id="pb"></i></div>
  <div style="margin-top:6px;font-size:13px;color:#555" id="at"></div>
</div>
<div class="k">
  <div><b id="tt">0</b>Total</div><div class="ok"><b id="ko">0</b>Removidos</div>
  <div class="pu"><b id="kp">0</b>Pulados (vazios)</div><div class="er"><b id="ke">0</b>Erros</div>
</div>
<div class="c" style="margin-top:14px"><div id="log"></div></div>
</div>
<script>
let arquivo=null,n=0,timer=null;
const $=id=>document.getElementById(id),drop=$('drop');
drop.onclick=()=>$('f').click();
$('f').onchange=e=>set(e.target.files[0]);
['dragover','dragenter'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add('on')}));
['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove('on')}));
drop.addEventListener('drop',e=>set(e.dataTransfer.files[0]));
function set(f){if(!f)return;arquivo=f;$('fn').textContent=f.name}
$('ini').onclick=async()=>{
  const fd=new FormData();if(arquivo)fd.append('csv',arquivo);
  fd.append('texto',$('t').value);fd.append('site',$('site').value);
  const r=await fetch('/iniciar',{method:'POST',body:fd}),j=await r.json();
  if(!r.ok){alert(j.erro);return}
  $('log').textContent='';n=0;poll();timer=setInterval(poll,700);
};
$('parar').onclick=()=>fetch('/parar',{method:'POST'});
$('rel').onclick=()=>location.href='/relatorio';
async function poll(){
  const j=await (await fetch('/status?desde='+n)).json();
  if(j.logs.length){const l=$('log');l.textContent+=j.logs.join('\\n')+'\\n';l.scrollTop=l.scrollHeight}
  n=j.n;$('tt').textContent=j.total;$('ko').textContent=j.ok;$('kp').textContent=j.pulado;$('ke').textContent=j.erro;
  $('pb').style.width=(j.total?100*j.feitos/j.total:0)+'%';
  $('at').textContent=j.rodando?(j.feitos+'/'+j.total+(j.atual?'  —  processando '+j.atual:'')):'';
  $('ini').disabled=j.rodando;$('parar').disabled=!j.rodando;$('rel').disabled=!j.tem_relatorio;
  if(!j.rodando&&timer){clearInterval(timer);timer=null}
}
poll();
</script></body></html>"""


if __name__ == "__main__":
    threading.Timer(1.0, lambda: webbrowser.open(f"http://localhost:{PORTA}")).start()
    app.run(host="127.0.0.1", port=PORTA, debug=False, threaded=True)

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Remove foto principal + galeria + arquivos de Download (Datasheet,
Desenho 3D, Manual) de produtos WooCommerce por SKU.

- NÃO apaga arquivos da Biblioteca de Mídia (só desvincula do produto).
- MODO RÁPIDO: antes de abrir a tela do produto, o script consulta a busca
  e a página de edição via HTTP (sem renderizar). Se o produto não tiver
  nada para remover, PULA na hora. Só abre o navegador nos que têm algo.
  Se a leitura rápida falhar por qualquer motivo, cai no modo antigo.

Os SKUs vêm de um CSV (ou .xlsx) colocado NA MESMA PASTA deste script:
  - Se existir "remover.csv", usa ele.
  - Senão, se tiver só 1 arquivo .csv na pasta, usa esse.
  - Ou passe o nome direto:  py remover_tudo_woocommerce.py minha_lista.csv

Uso:
  py remover_tudo_woocommerce.py
"""

import csv
import html as htmllib
import io
import re
import sys
import time
from html.parser import HTMLParser
from pathlib import Path

from playwright.sync_api import sync_playwright, Page, FrameLocator

BASE_URL = "https://sagloja.devsplcomunic.com.br"
SCRIPT_DIR = Path(__file__).resolve().parent
PROFILE_DIR = SCRIPT_DIR / "chrome-profile-sag"
CSV_PADRAO = "remover.csv"
COLUNAS_SKU = ("sku", "cigam", "codigo", "código", "referencia", "referência", "cod")
REPORT_FILE = SCRIPT_DIR / "resultado_remocao_tudo_woocommerce.txt"

TIMEOUT = 30000


def log(msg: str):
    print(msg, flush=True)


# ---------------------------------------------------------------- CSV

def acha_csv() -> Path:
    if len(sys.argv) > 1:
        p = Path(sys.argv[1])
        if not p.is_absolute():
            p = SCRIPT_DIR / p
        if not p.exists():
            raise FileNotFoundError(f"CSV não encontrado: {p}")
        return p

    padrao = SCRIPT_DIR / CSV_PADRAO
    if padrao.exists():
        return padrao

    csvs = sorted(SCRIPT_DIR.glob("*.csv"))
    if len(csvs) == 1:
        return csvs[0]
    if not csvs:
        raise FileNotFoundError(
            f"Nenhum CSV na pasta {SCRIPT_DIR}. Coloque a lista de SKUs como '{CSV_PADRAO}'."
        )
    nomes = ", ".join(c.name for c in csvs)
    raise FileNotFoundError(
        f"Tem mais de um CSV na pasta ({nomes}). Renomeie o certo para '{CSV_PADRAO}' "
        "ou passe o nome: py remover_tudo_woocommerce.py arquivo.csv"
    )


def _limpa_sku(v) -> str:
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    v = ("" if v is None else str(v)).strip().strip('"').strip()
    if re.fullmatch(r"\d+\.0+", v):
        v = v.split(".")[0]
    return v


def _linhas_do_xlsx(raw: bytes) -> list[list[str]]:
    from openpyxl import load_workbook
    ws = load_workbook(io.BytesIO(raw), data_only=True).active
    return [[_limpa_sku(c) for c in row] for row in ws.iter_rows(values_only=True)]


def _linhas_do_csv(raw: bytes) -> list[list[str]]:
    texto = None
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            texto = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if texto is None:
        raise ValueError("Não consegui ler a codificação do CSV.")

    linhas = [l for l in texto.splitlines() if l.strip()]
    if not linhas:
        return []

    if not any(d in linhas[0] for d in ",;\t"):
        return [[l] for l in linhas]

    try:
        delim = csv.Sniffer().sniff(linhas[0], delimiters=",;\t").delimiter
    except csv.Error:
        delim = ";" if ";" in linhas[0] else ","
    return list(csv.reader(linhas, delimiter=delim))


def carrega_skus(caminho) -> list[str]:
    """Lê os SKUs de um .csv ou .xlsx/.xlsm. Aceita caminho (str ou Path)
    ou um arquivo já aberto (ex.: upload do Flask)."""
    if hasattr(caminho, "read"):
        nome = getattr(caminho, "filename", None) or getattr(caminho, "name", "") or ""
        raw = caminho.read()
        if isinstance(raw, str):
            raw = raw.encode("utf-8")
    else:
        caminho = Path(caminho)
        nome = caminho.name
        raw = caminho.read_bytes()

    ext = Path(str(nome)).suffix.lower()
    if ext in (".xlsx", ".xlsm") or raw[:2] == b"PK":   # xlsx é um zip (começa com PK)
        rows = _linhas_do_xlsx(raw)
    else:
        rows = _linhas_do_csv(raw)

    rows = [r for r in rows if any((c or "").strip() for c in r)]
    if not rows:
        return []

    cab = [(c or "").strip().lower() for c in rows[0]]
    idx = next((cab.index(c) for c in COLUNAS_SKU if c in cab), None)
    if idx is not None:
        dados = rows[1:]
    else:
        idx = 0
        dados = rows[1:] if not re.search(r"\d", rows[0][0] if rows[0] else "") else rows

    skus, vistos = [], set()
    for r in dados:
        if idx >= len(r):
            continue
        sku = _limpa_sku(r[idx])
        if not sku:
            continue
        if re.search(r"[eE]\+", sku):
            log(f"  AVISO: SKU '{sku}' veio em notação científica (Excel) — corrija no CSV. Ignorado.")
            continue
        if sku not in vistos:
            vistos.add(sku)
            skus.append(sku)
    return skus


# ---------------------------------------------------------------- MODO RÁPIDO (HTTP)

class _Inputs(HTMLParser):
    def __init__(self):
        super().__init__()
        self.inputs = []

    def handle_starttag(self, tag, attrs):
        if tag == "input":
            self.inputs.append(dict(attrs))

    handle_startendtag = handle_starttag


def _get_html(page: Page, url: str) -> str | None:
    """Baixa a página usando os cookies do Chrome logado, sem renderizar."""
    try:
        resp = page.context.request.get(url, timeout=TIMEOUT)
        if not resp.ok:
            return None
        return resp.text()
    except Exception:
        return None


def ler_estado_rapido(page: Page, post_id: int):
    """Lê foto principal / galeria / downloads direto do HTML da edição.
    Retorna dict ou None se não conseguir ler."""
    html = _get_html(page, f"{BASE_URL}/wp-admin/post.php?post={post_id}&action=edit")
    if not html or "product_image_gallery" not in html:
        return None

    p = _Inputs()
    p.feed(html)

    thumb = gal = None
    docs = 0
    for a in p.inputs:
        nome = a.get("name") or ""
        ident = a.get("id") or ""
        val = (a.get("value") or "").strip()
        if ident == "_thumbnail_id" or nome == "_thumbnail_id":
            thumb = val
        elif ident == "product_image_gallery" or nome == "product_image_gallery":
            gal = val
        elif nome.startswith("wcda_arquivos") and nome.endswith("[label]") and val:
            docs += 1

    if gal is None:
        return None

    return {
        "principal": thumb not in (None, "", "-1", "0"),
        "galeria": len([x for x in gal.split(",") if x.strip()]),
        "downloads": docs,
    }


# ---------------------------------------------------------------- NAVEGADOR

def admin_frame(page: Page) -> FrameLocator:
    iframe = page.locator("#frame-adminify-app--iframe")
    iframe.wait_for(state="attached", timeout=TIMEOUT)
    return iframe.content_frame


def wait_for_login(page: Page):
    log("Verificando login...")
    page.goto(f"{BASE_URL}/wp-admin/edit.php?post_type=product", wait_until="domcontentloaded", timeout=60000)

    if "wp-login.php" in page.url or page.locator("#loginform").count() > 0:
        log("\n>>> FAÇA LOGIN MANUALMENTE NA JANELA DO CHROME. <<<")
        log("O script continua sozinho assim que detectar o login.\n")
        while True:
            time.sleep(1.5)
            try:
                if "wp-login.php" not in page.url and page.locator("#loginform").count() == 0:
                    break
            except Exception:
                pass
        log("Login detectado.")
        page.goto(f"{BASE_URL}/wp-admin/edit.php?post_type=product", wait_until="domcontentloaded", timeout=60000)


def open_product_by_id(page: Page, post_id: int):
    page.goto(f"{BASE_URL}/wp-admin/post.php?post={post_id}&action=edit",
              wait_until="domcontentloaded", timeout=60000)
    frame = admin_frame(page)
    frame.locator("#title").wait_for(state="visible", timeout=TIMEOUT)


def open_product_by_sku(page: Page, sku: str) -> bool:
    """Modo antigo (fallback): busca pela tela."""
    page.goto(f"{BASE_URL}/wp-admin/edit.php?post_type=product", wait_until="domcontentloaded", timeout=60000)
    frame = admin_frame(page)

    search_box = frame.get_by_role("searchbox", name="Pesquisar produtos:")
    search_box.wait_for(state="visible", timeout=TIMEOUT)
    search_box.fill(sku)
    frame.get_by_role("button", name="Pesquisar produtos").click()
    page.wait_for_timeout(1000)

    sku_pattern = re.compile(rf"^\s*{re.escape(sku)}\s*$")
    row = None
    for _ in range(30):
        try:
            sku_cell = frame.locator("td.column-sku").filter(has_text=sku_pattern)
            if sku_cell.count() > 0:
                row = sku_cell.first.locator("xpath=ancestor::tr")
                break
        except Exception:
            pass
        page.wait_for_timeout(500)

    if row is None:
        return False

    link = row.locator("a.row-title")
    if link.count() == 0:
        link = row.locator("strong a")
    if link.count() == 0:
        link = row.locator("a").first

    href = link.get_attribute("href")
    if href:
        page.goto(href, wait_until="domcontentloaded", timeout=60000)
    else:
        link.click()

    frame = admin_frame(page)
    frame.locator("#title").wait_for(state="visible", timeout=TIMEOUT)
    return True


def featured_image_is_set(frame: FrameLocator) -> bool:
    try:
        thumb = frame.locator("#_thumbnail_id")
        if thumb.count() > 0:
            val = (thumb.first.input_value(timeout=3000) or "").strip()
            return val not in ("", "-1", "0")
    except Exception:
        pass
    try:
        return frame.locator("#set-post-thumbnail img").count() > 0
    except Exception:
        return False


def gallery_count(frame: FrameLocator) -> int:
    try:
        val = frame.locator("#product_image_gallery").input_value(timeout=3000)
        return len([x for x in val.split(",") if x.strip()])
    except Exception:
        return -1


def remove_featured_image(frame: FrameLocator):
    if not featured_image_is_set(frame):
        return False

    try:
        remove = frame.locator("#remove-post-thumbnail")
        if remove.count() > 0:
            try:
                remove.first.click(timeout=5000, force=True)
            except Exception:
                pass
    except Exception:
        pass

    resultado = frame.locator("body").evaluate(
        """() => {
            const input =
                document.querySelector('#_thumbnail_id') ||
                document.querySelector('input[name="_thumbnail_id"]');
            if (!input) return {ok:false, motivo:'campo _thumbnail_id não encontrado'};
            input.value = '-1';
            input.setAttribute('value', '-1');
            input.dispatchEvent(new Event('input', {bubbles:true}));
            input.dispatchEvent(new Event('change', {bubbles:true}));
            const box = document.querySelector('#set-post-thumbnail');
            if (box) { const img = box.querySelector('img'); if (img) img.remove(); }
            return {ok:true, valor:input.value};
        }"""
    )

    if not resultado or not resultado.get("ok"):
        motivo = resultado.get("motivo") if isinstance(resultado, dict) else "desconhecido"
        raise RuntimeError(f"Não consegui limpar a imagem principal: {motivo}")
    return True


def clear_gallery(frame: FrameLocator):
    gal = frame.locator("#product_image_gallery")
    if gal.count() == 0 or gallery_count(frame) == 0:
        return False

    gal.evaluate(
        """el => {
            el.value = '';
            el.setAttribute('value', '');
            el.dispatchEvent(new Event('input', {bubbles:true}));
            el.dispatchEvent(new Event('change', {bubbles:true}));
            const box = document.querySelector('#product_images_container');
            if (box) box.querySelectorAll('ul.product_images li.image').forEach(li => li.remove());
        }"""
    )
    return True


def download_files_count(frame: FrameLocator) -> int:
    try:
        return frame.locator("body").evaluate(
            """() => {
                const labels = document.querySelectorAll('input[name^="wcda_arquivos"][name$="[label]"]');
                let n = 0;
                labels.forEach(el => { if (el.value && el.value.trim()) n++; });
                return n;
            }"""
        )
    except Exception:
        return -1


def clear_download_files(frame: FrameLocator) -> bool:
    if download_files_count(frame) <= 0:
        return False

    frame.locator("body").evaluate(
        """() => {
            document.querySelectorAll('input[name^="wcda_arquivos"]').forEach(el => {
                if (el.value) {
                    el.value = '';
                    el.setAttribute('value', '');
                    el.dispatchEvent(new Event('input', {bubbles:true}));
                    el.dispatchEvent(new Event('change', {bubbles:true}));
                }
            });
        }"""
    )
    return True


def click_update(page: Page, frame: FrameLocator):
    try:
        frame.locator("body").evaluate("() => window.scrollTo(0, 0)")
    except Exception:
        pass
    page.wait_for_timeout(300)

    btn = frame.locator("#publish")
    if btn.count() == 0:
        btn = frame.get_by_role("button", name=re.compile(r"^Atualizar$", re.I))

    btn.first.wait_for(state="visible", timeout=TIMEOUT)
    try:
        btn.first.click(timeout=10000)
    except Exception:
        try:
            btn.first.click(timeout=10000, force=True)
        except Exception:
            btn.first.evaluate("el => el.click()")

    page.wait_for_timeout(2500)
    try:
        admin_frame(page).locator("#title").wait_for(state="visible", timeout=TIMEOUT)
    except Exception:
        pass


def verify_clean(page: Page) -> tuple[bool, str]:
    frame = admin_frame(page)
    problems = []
    if featured_image_is_set(frame):
        problems.append("imagem principal ainda definida")
    gal = gallery_count(frame)
    if gal > 0:
        problems.append(f"galeria ainda tem {gal} imagem(ns)")
    docs = download_files_count(frame)
    if docs > 0:
        problems.append(f"ainda tem {docs} arquivo(s) de download preenchido(s)")
    return (not problems, "; ".join(problems))


# ---------------------------------------------------------------- FLUXO

def buscar_id_pela_tela(page: Page, sku: str):
    """Pesquisa o SKU na barra de 'Pesquisar produtos' e devolve o ID do post
    (o número do post.php?post=XXXX). Retorna None se não achar."""
    page.goto(f"{BASE_URL}/wp-admin/edit.php?post_type=product",
              wait_until="domcontentloaded", timeout=60000)
    frame = admin_frame(page)

    box = frame.locator("#post-search-input")
    box.wait_for(state="visible", timeout=TIMEOUT)
    box.fill(sku)
    frame.locator("#search-submit").click()

    sku_pattern = re.compile(rf"^\s*{re.escape(sku)}\s*$")
    for _ in range(30):
        try:
            cell = frame.locator("td.column-sku").filter(has_text=sku_pattern)
            if cell.count() > 0:
                row = cell.first.locator("xpath=ancestor::tr")
                tr_id = row.get_attribute("id") or ""          # "post-6396"
                m = re.search(r"post-(\d+)", tr_id)
                if not m:
                    href = row.locator("a.row-title").first.get_attribute("href") or ""
                    m = re.search(r"post=(\d+)", href)
                if m:
                    return int(m.group(1))
        except Exception:
            pass
        page.wait_for_timeout(300)
    return None


def process_sku(page: Page, sku: str) -> str:
    """Retorna 'OK' ou 'PULADO'. Lança exceção em erro."""
    post_id = buscar_id_pela_tela(page, sku)
    if not post_id:
        raise RuntimeError("Produto não encontrado pelo SKU")
    log(f"  ID {post_id}")

    estado = ler_estado_rapido(page, post_id)
    if estado is not None:
        if not estado["principal"] and estado["galeria"] == 0 and estado["downloads"] == 0:
            log("  PULADO - produto já está vazio.")
            return "PULADO"
        log(
            f"  Tem algo: principal={'SIM' if estado['principal'] else 'NÃO'} | "
            f"galeria={estado['galeria']} | downloads={estado['downloads']}"
        )

    open_product_by_id(page, post_id)
    frame = admin_frame(page)

    changed = False
    if remove_featured_image(frame):
        log("  Foto principal removida.")
        changed = True
    if clear_gallery(frame):
        log("  Galeria limpa.")
        changed = True
    if clear_download_files(frame):
        log("  Downloads (Datasheet/Desenho 3D/Manual) limpos.")
        changed = True

    if not changed:
        log("  PULADO - nada para remover.")
        return "PULADO"

    click_update(page, frame)
    ok, detail = verify_clean(page)
    if not ok:
        raise RuntimeError("Depois de salvar: " + detail)
    log("  OK - salvo e confirmado.")
    return "OK"


def main():
    csv_path = acha_csv()
    skus = carrega_skus(csv_path)
    log(f"CSV: {csv_path.name} -> {len(skus)} SKU(s)")
    if not skus:
        log("Nenhum SKU encontrado no CSV. Nada a fazer.")
        return
    log(f"IMPORTANTE: feche outros scripts/Chromes usando o mesmo {PROFILE_DIR.name}.\n")

    resultados = []

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            viewport={"width": 1440, "height": 900},
        )
        page = context.pages[0] if context.pages else context.new_page()
        wait_for_login(page)

        total = len(skus)
        for i, sku in enumerate(skus, start=1):
            log(f"\n[{i}/{total}] SKU {sku}")
            try:
                status = process_sku(page, sku)
                resultados.append((sku, status, ""))
            except Exception as e:
                msg = str(e)
                log(f"  !! ERRO: {msg}")
                resultados.append((sku, "ERRO", msg))
                try:
                    debug_dir = SCRIPT_DIR / "debug_remocao_fotos"
                    debug_dir.mkdir(exist_ok=True)
                    page.screenshot(path=str(debug_dir / f"erro_{sku}.png"), full_page=True)
                except Exception:
                    pass

        context.close()

    with REPORT_FILE.open("w", encoding="utf-8") as f:
        for sku, status, detail in resultados:
            f.write(f"{sku}\t{status}\t{detail}\n")

    ok = sum(1 for _, s, _ in resultados if s == "OK")
    pulado = sum(1 for _, s, _ in resultados if s == "PULADO")
    erro = sum(1 for _, s, _ in resultados if s == "ERRO")
    log("\n=== FINALIZADO ===")
    log(f"OK: {ok} | PULADOS: {pulado} | ERROS: {erro} | TOTAL: {len(resultados)}")
    log(f"Relatório: {REPORT_FILE}")


if __name__ == "__main__":
    main()

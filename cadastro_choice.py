"""
Automação de Cadastro de Produtos - CHOICE (WooCommerce)
==========================================================

O que este script faz, para cada SKU digitado:
  1. Procura a pasta do produto dentro de ROOT_FOLDER (busca recursiva pelo
     nome exato da pasta = SKU, igual ao que você faz na busca do Explorer).
  2. Dentro da pasta, identifica os arquivos pelo padrão de nome:
       *_variante.*   -> imagem principal (foto com escrita da variante)
       *_foto1.. N    -> fotos da galeria (todas, em ordem numérica)
       *_datasheet.*  -> PDF do datasheet
       *_3d.zip       -> desenho 3D zipado (opcional - só usa se existir)
  3. Abre o produto no wp-admin pelo SKU.
  4. Define a imagem principal (variante).
  5. Adiciona as fotos na galeria.
  6. Cria o campo "Datasheet" em Download de Arquivos e sobe o PDF.
  7. Se existir, cria o campo "Desenho 3D" e sobe o ZIP.
  8. Clica em Atualizar.

Depois de terminar um SKU, ele pergunta o próximo. Digite "sair" para terminar.

IMPORTANTE - por que esse script existe do jeito que é:
  Todo o conteúdo do wp-admin nesse site roda dentro de um <iframe> do
  plugin Adminify (id "frame-adminify-app--iframe"). Por isso, toda
  interação passa por admin_frame(page) em vez de "page" direto.

-----------------------------------------------------------------------
COMO USAR (Windows)
-----------------------------------------------------------------------
1. Instalar dependências (uma vez só):
       py -m pip install playwright openpyxl
       py -m playwright install chromium

2. Ajustar ROOT_FOLDER abaixo se a letra do drive de rede for diferente
   no seu PC.

3. Rodar:
       py cadastro_choice.py

   Na primeira vez, uma janela do Chrome vai abrir. Faça login manualmente
   no wp-admin da Choice nessa janela, volte no terminal e aperte ENTER.
   Da próxima vez que rodar, ele reaproveita o login salvo.

4. Escolha entre processar uma planilha (.xlsx/.csv com uma coluna de
   SKU/Código) de uma vez, ou digitar os SKUs um por um. No final do lote
   sai um resumo com o que deu certo e o que falhou, e um arquivo
   "resultado_lote_*.txt" é salvo na pasta com o mesmo resumo.

5. Teste com 1-2 produtos antes de rodar o lote inteiro. Se algum passo
   falhar, me manda o erro do terminal (e se der, um novo trecho gravado
   com "py -m playwright codegen <url>" fazendo só aquele pedaço que
   travou - isso é ouro pra eu corrigir rápido).

-----------------------------------------------------------------------
"""

import re
import time
from pathlib import Path

from playwright.sync_api import sync_playwright, Page, FrameLocator, Locator


# ======================= CONFIG (ajuste aqui) =======================

BASE_URL = "https://choice.devsplcomunic.com.br"

# Pasta raiz onde ficam as pastas dos produtos (cada pasta = nome do SKU)
ROOT_FOLDER = r"Z:\Choice\Geral\E-Commerce\CHOICE\Automação"

# Onde o navegador guarda o login salvo entre execuções.
PROFILE_DIR = r".\chrome-profile-choice"

# Tempo máximo de espera (ms) para uploads/telas carregarem
UPLOAD_TIMEOUT = 90000

# ======================================================================


# --------------------------- arquivos/pastas ---------------------------

def find_product_folder(root: Path, sku: str) -> Path | None:
    """Procura recursivamente por uma pasta cujo nome seja exatamente o SKU."""
    sku_lower = sku.strip().lower()
    if not root.exists():
        print(f"  !! Pasta raiz não existe ou não está acessível: {root}")
        return None
    for path in root.rglob("*"):
        if path.is_dir() and path.name.lower() == sku_lower:
            return path
    return None


def _classify_files_padrao_choice(folder: Path) -> dict:
    """Padrão antigo (fornecedor Choice): arquivos soltos na pasta, nomeados
    ..._variante, ..._fotoN, ..._datasheet, ..._3d.zip"""
    files = [f for f in folder.iterdir() if f.is_file()]

    def find_one(suffix_regex: str, ext: str | None = None):
        for f in files:
            if re.search(suffix_regex, f.stem, re.IGNORECASE):
                if ext and f.suffix.lower() != ext:
                    continue
                return f
        return None

    variante = find_one(r"_variante$")
    datasheet = find_one(r"_datasheet$")
    desenho3d = find_one(r"_3d$", ext=".zip")

    fotos = []
    for f in files:
        m = re.search(r"_foto(\d+)$", f.stem, re.IGNORECASE)
        if m:
            fotos.append((int(m.group(1)), f))
    fotos.sort(key=lambda t: t[0])
    fotos = [f for _, f in fotos]

    return {"variante": variante, "fotos": fotos, "datasheet": datasheet, "desenho3d": desenho3d}


def _classify_files_padrao_subpastas(folder: Path) -> dict:
    """Padrão novo (fornecedor com subpastas): Fotos/ (a que termina em
    FRONT é a principal, as demais vão pra galeria), e os arquivos de
    Datasheet e Desenho 3D são os .zip prontos na raiz da pasta do produto
    (Datasheets.zip / Desenho 3D.zip) - mesma lógica de "vem zipado" que
    já era usada pro desenho 3D do outro fornecedor."""
    fotos_dir = folder / "Fotos"

    variante = None
    fotos = []
    if fotos_dir.is_dir():
        imagens = sorted(
            f for f in fotos_dir.iterdir()
            if f.is_file() and f.suffix.lower() in (".png", ".jpg", ".jpeg", ".webp")
        )
        for f in imagens:
            if re.search(r"front$", f.stem, re.IGNORECASE):
                fotos.append(f)
            else:
                variante = f

    def zip_na_raiz(nome: str) -> Path | None:
        candidato = folder / f"{nome}.zip"
        return candidato if candidato.is_file() else None

    datasheet = zip_na_raiz("Datasheets")
    desenho3d = zip_na_raiz("Desenho 3D")

    return {"variante": variante, "fotos": fotos, "datasheet": datasheet, "desenho3d": desenho3d}


def classify_files(folder: Path) -> dict:
    """Detecta automaticamente qual padrão de pasta é esse (fornecedor
    antigo com arquivos soltos, ou fornecedor novo com subpasta Fotos/) e
    usa a lógica certa pra cada um."""
    if (folder / "Fotos").is_dir():
        return _classify_files_padrao_subpastas(folder)
    return _classify_files_padrao_choice(folder)


# --------------------------- navegador / iframe ---------------------------

def admin_frame(page: Page) -> FrameLocator:
    """Todo o conteúdo do wp-admin nesse site fica dentro de um iframe do
    plugin Adminify. É por isso que quase toda interação passa por aqui em
    vez de usar 'page' direto."""
    return page.locator("#frame-adminify-app--iframe").content_frame


def ensure_logged_in(page: Page, base_url: str):
    # A tela de login em si NÃO fica dentro do iframe do Adminify.
    page.goto(f"{base_url}/wp-admin/edit.php?post_type=product")
    if "wp-login.php" in page.url or page.locator("#loginform").count() > 0:
        print("\n>>> Faça login manualmente na janela do Chrome que abriu.")
        input(">>> Depois de logar, volte aqui e pressione ENTER para continuar... ")
        page.goto(f"{base_url}/wp-admin/edit.php?post_type=product")


def clear_stale_selections(frame: FrameLocator, max_clicks: int = 25):
    """Desmarca qualquer item que tenha ficado selecionado de uma ação
    anterior. A biblioteca de mídia mantém seleção entre uma abertura e
    outra do modal, o que pode acabar juntando o arquivo errado."""
    for _ in range(max_clicks):
        btn = frame.get_by_role("button", name="Desmarcar").first
        try:
            if btn.count() == 0:
                return
            btn.click(timeout=2000)
        except Exception:
            return


def wait_enabled_and_click(locator: Locator, timeout_ms: int = UPLOAD_TIMEOUT):
    """Espera o botão ficar clicável (sem disabled/aria-disabled/pointer-
    events:none) e clica."""
    start = time.time()
    while (time.time() - start) * 1000 < timeout_ms:
        try:
            ready = locator.evaluate(
                """el => {
                    if (el.disabled) return false;
                    if (el.getAttribute('aria-disabled') === 'true') return false;
                    const s = window.getComputedStyle(el);
                    if (s.pointerEvents === 'none') return false;
                    return true;
                }"""
            )
        except Exception:
            ready = False
        if ready:
            break
        time.sleep(0.3)
    locator.click(timeout=15000)


def upload_and_confirm(page: Page, frame: FrameLocator, files: list[Path], confirm_button_name: str):
    """Fluxo padrão dentro de um modal de mídia já aberto: limpa seleções
    antigas, garante a aba 'Enviar arquivos', sobe o(s) arquivo(s) e clica
    no botão de confirmação (ex: 'Definir imagem do produto',
    'Adicionar na galeria', 'Usar este arquivo').

    O botão "Selecionar arquivos" é um <button> comum que só aciona um
    campo de arquivo escondido por trás - por isso interceptamos o
    "seletor de arquivo" do navegador em vez de jogar o arquivo direto
    no botão."""
    clear_stale_selections(frame)

    try:
        frame.get_by_role("tab", name="Enviar arquivos").click(timeout=5000)
    except Exception:
        pass  # já pode estar nessa aba

    browse = frame.get_by_role("button", name="Selecionar arquivos")
    with page.expect_file_chooser(timeout=15000) as fc_info:
        browse.click()
    fc_info.value.set_files([str(f) for f in files])

    confirm = frame.get_by_role("button", name=confirm_button_name)
    confirm.wait_for(state="visible", timeout=UPLOAD_TIMEOUT)

    # dá um respiro pro upload/seleção automática terminar antes de confirmar
    page.wait_for_timeout(800)

    wait_enabled_and_click(confirm, timeout_ms=UPLOAD_TIMEOUT)


def open_product_by_sku(page: Page, sku: str, base_url: str) -> bool:
    page.goto(f"{base_url}/wp-admin/edit.php?post_type=product")
    frame = admin_frame(page)

    search_box = frame.get_by_role("searchbox", name="Pesquisar produtos:")
    search_box.wait_for(state="visible", timeout=20000)
    search_box.click()
    search_box.fill(sku)
    frame.get_by_role("button", name="Pesquisar produtos").click()

    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        pass

    # Em vez de confiar que a busca já filtrou a lista (a lista pode ainda
    # não ter atualizado no momento em que a gente olha), procuramos direto
    # pela linha cuja coluna SKU seja EXATAMENTE igual ao que foi pedido.
    # Isso evita abrir o produto errado se a busca demorar ou não filtrar.
    sku_pattern = re.compile(rf"^\s*{re.escape(sku)}\s*$")
    row = None
    for _ in range(24):  # ~12s de tentativas
        sku_cell = frame.locator("td.column-sku").filter(has_text=sku_pattern)
        try:
            if sku_cell.count() > 0:
                row = sku_cell.first.locator("xpath=ancestor::tr")
                break
        except Exception:
            pass
        page.wait_for_timeout(500)

    if row is None:
        print(f"  !! Não encontrei nenhuma linha com o SKU exatamente igual a '{sku}' na lista.")
        return False

    title_link = row.locator("a.row-title")
    if title_link.count() == 0:
        title_link = row.locator("strong a")
    if title_link.count() == 0:
        title_link = row.locator("a").first

    nome_produto = title_link.inner_text().strip()
    href = title_link.get_attribute("href")
    print(f"  Abrindo produto: {nome_produto}")

    if href:
        page.goto(href)
    else:
        title_link.click()

    frame = admin_frame(page)
    try:
        frame.locator("#title").wait_for(state="visible", timeout=20000)
    except Exception:
        print(f"  !! Não consegui confirmar que abriu a página de edição. URL atual: {page.url}")
        return False

    return True


def featured_image_is_set(frame: FrameLocator) -> bool:
    """Confere se a imagem principal realmente está definida (mostra uma
    miniatura de verdade, não só o texto/link de 'definir')."""
    try:
        return frame.locator("#set-post-thumbnail img").count() > 0
    except Exception:
        return False


def gallery_image_count(frame: FrameLocator) -> int:
    """Conta quantas imagens estão na galeria pelo campo escondido que o
    WooCommerce usa por baixo dos panos. Retorna -1 se não conseguir
    checar (não trava o processo, só não confirma)."""
    try:
        val = frame.locator("#product_image_gallery").input_value(timeout=3000)
        return len([x for x in val.split(",") if x.strip()])
    except Exception:
        return -1


def set_featured_image(page: Page, frame: FrameLocator, path: Path | None):
    if not path:
        print("  ! Aviso: arquivo '_variante' não encontrado - pulando imagem principal.")
        return

    for tentativa in (1, 2):
        print(f"  Definindo imagem principal: {path.name} (tentativa {tentativa})")
        frame.locator("#set-post-thumbnail").click(timeout=15000)
        upload_and_confirm(page, frame, [path], confirm_button_name="Definir imagem do produto")
        page.wait_for_timeout(600)
        if featured_image_is_set(frame):
            print("  Imagem principal confirmada na tela.")
            return
        print("  ! A imagem principal não apareceu confirmada - tentando de novo...")

    raise RuntimeError("Imagem principal não confirmou depois de 2 tentativas.")


def add_gallery_images(page: Page, frame: FrameLocator, fotos: list[Path]):
    if not fotos:
        print("  ! Aviso: nenhuma foto '_fotoN' encontrada - pulando galeria.")
        return

    esperado = len(fotos)
    for tentativa in (1, 2):
        print(f"  Adicionando {esperado} foto(s) na galeria (tentativa {tentativa}).")
        frame.get_by_role("link", name="Adicionar galeria de imagens").click()
        upload_and_confirm(page, frame, fotos, confirm_button_name="Adicionar na galeria")
        page.wait_for_timeout(600)

        atual = gallery_image_count(frame)
        if atual == -1:
            print("  ! Não consegui confirmar a quantidade de fotos na galeria (seguindo mesmo assim).")
            return
        if atual >= esperado:
            print(f"  Galeria confirmada com {atual} foto(s) na tela.")
            return
        print(f"  ! Galeria mostrou {atual} foto(s), esperava pelo menos {esperado} - tentando de novo...")

    raise RuntimeError(f"Galeria não atingiu o número esperado de fotos ({esperado}) depois de 2 tentativas.")


def add_download_file(page: Page, frame: FrameLocator, nome: str, path: Path):
    print(f"  Adicionando arquivo de download '{nome}': {path.name}")

    label_inputs = frame.locator('input[name^="wcda_arquivos"][name$="[label]"]')
    index = label_inputs.count()

    frame.get_by_role("button", name="+ Adicionar Arquivo").click()
    label_input = label_inputs.nth(index)
    label_input.fill(nome)

    for tentativa in (1, 2):
        print(f"  Selecionando arquivo para '{nome}' (tentativa {tentativa})...")
        frame.get_by_role("button", name="Selecionar").nth(index).click()
        upload_and_confirm(page, frame, [path], confirm_button_name="Usar este arquivo")

        if _arquivo_download_confirmado(frame, index):
            print(f"  Arquivo '{nome}' confirmado (campo preenchido).")
            return
        print(f"  ! Não vi o campo do arquivo '{nome}' preenchido - tentando de novo...")

    raise RuntimeError(f"Não consegui confirmar que o arquivo '{nome}' foi anexado, mesmo após tentar de novo.")


def _arquivo_download_confirmado(frame: FrameLocator, index: int, timeout_ms: int = 5000) -> bool:
    """Confirma se o arquivo da linha 'index' foi vinculado de verdade,
    checando pelo NOME EXATO do campo (wcda_arquivos[index][...]) se algum
    campo dessa linha, além do de nome/label, recebeu um valor. Isso não
    depende de como a linha está organizada visualmente na página - usa o
    próprio atributo name do campo, que é estável.

    Faz um retry curto (poll) porque o campo pode demorar um instante pra
    ser preenchido depois do upload."""
    start = time.time()
    while (time.time() - start) * 1000 < timeout_ms:
        try:
            inputs = frame.locator(f'input[name^="wcda_arquivos[{index}]"]')
            for i in range(inputs.count()):
                campo = inputs.nth(i)
                nome_campo = campo.get_attribute("name") or ""
                if nome_campo.endswith("[label]"):
                    continue
                try:
                    valor = campo.input_value(timeout=1000)
                except Exception:
                    continue
                if valor and valor.strip():
                    return True
        except Exception:
            pass
        time.sleep(0.4)
    return False


def save_product(
    page: Page,
    frame: FrameLocator,
    esperado_imagem: bool,
    esperado_fotos: int,
    esperado_labels: list[str],
):
    print("  Salvando produto (Atualizar)...")
    frame.get_by_role("button", name="Atualizar").click(timeout=15000)

    # Esse clique dispara um envio de formulário de verdade (é isso que
    # realmente salva a imagem/galeria/arquivos no banco). Se a gente for
    # pro próximo produto rápido demais, corta esse envio pela metade e o
    # produto fica sem salvar o que tinha sido adicionado.
    #
    # 1) Espera mínima fixa de segurança - o envio pro servidor não é
    #    instantâneo, isso garante uma margem mesmo se as esperas abaixo
    #    falharem silenciosamente.
    page.wait_for_timeout(2500)

    # 2) Tenta esperar de verdade o iframe (não a página de fora) terminar
    #    de recarregar - pegando o Frame real por trás do FrameLocator.
    try:
        handle = page.locator("#frame-adminify-app--iframe").element_handle(timeout=5000)
        real_frame = handle.content_frame() if handle else None
    except Exception:
        real_frame = None

    if real_frame is not None:
        try:
            real_frame.wait_for_load_state("networkidle", timeout=25000)
        except Exception:
            pass
    else:
        try:
            page.wait_for_load_state("networkidle", timeout=25000)
        except Exception:
            pass

    # 3) Confirmação visual final: o título do produto reapareceu na tela?
    frame_apos = admin_frame(page)
    try:
        frame_apos.locator("#title").wait_for(state="visible", timeout=20000)
        print("  Página recarregou depois de salvar - confirmado.")
    except Exception:
        raise RuntimeError("A página não confirmou visualmente que recarregou depois de salvar.")

    # Checagem final: depois do recarregamento (ou seja, depois de já ter
    # sido salvo no banco), a imagem/galeria/arquivos realmente
    # persistiram? Se não, isso conta como ERRO nesse SKU - não segue
    # como se tivesse dado certo.
    problemas = []

    if esperado_imagem and not featured_image_is_set(frame_apos):
        problemas.append("imagem principal não ficou definida")

    if esperado_fotos > 0:
        atual = gallery_image_count(frame_apos)
        if atual != -1 and atual < esperado_fotos:
            problemas.append(f"galeria ficou com {atual} foto(s), esperava {esperado_fotos}")

    if esperado_labels:
        label_inputs_apos = frame_apos.locator('input[name^="wcda_arquivos"][name$="[label]"]')
        valores_atuais = []
        for i in range(label_inputs_apos.count()):
            try:
                valores_atuais.append(label_inputs_apos.nth(i).input_value())
            except Exception:
                pass
        faltando = [l for l in esperado_labels if l not in valores_atuais]
        if faltando:
            problemas.append(f"campo(s) de download não persistiram: {', '.join(faltando)}")

    if problemas:
        raise RuntimeError("Depois de salvar: " + "; ".join(problemas))


def save_debug_screenshot(page: Page, sku: str) -> Path | None:
    """Salva um print da tela no momento do erro, pra facilitar o diagnóstico."""
    try:
        debug_dir = Path("./debug_screenshots")
        debug_dir.mkdir(exist_ok=True)
        ts = time.strftime("%Y%m%d-%H%M%S")
        path = debug_dir / f"erro_{sku}_{ts}.png"
        page.screenshot(path=str(path), full_page=True)
        return path
    except Exception:
        return None


def process_sku(page: Page, sku: str, root: Path, base_url: str):
    print(f"\n--- Processando SKU {sku} ---")

    folder = find_product_folder(root, sku)
    if not folder:
        raise RuntimeError(f"Pasta não encontrada em {root} para o SKU {sku}.")
    print(f"  Pasta encontrada: {folder}")

    files = classify_files(folder)
    print(
        "  Arquivos identificados: "
        f"variante={files['variante'].name if files['variante'] else 'NENHUM'} | "
        f"fotos={[f.name for f in files['fotos']] if files['fotos'] else 'NENHUMA'} | "
        f"datasheet={files['datasheet'].name if files['datasheet'] else 'NENHUM'} | "
        f"desenho3d={files['desenho3d'].name if files['desenho3d'] else 'NENHUM'}"
    )
    if not any([files["variante"], files["fotos"], files["datasheet"], files["desenho3d"]]):
        print(
            "  ! ATENÇÃO: não identifiquei NENHUM arquivo reconhecível nessa pasta. "
            "Confira se a estrutura bate com o padrão esperado (arquivos "
            "_variante/_fotoN/_datasheet/_3d.zip OU subpasta 'Fotos' + "
            "'Datasheets.zip'/'Desenho 3D.zip')."
        )

    if not open_product_by_sku(page, sku, base_url):
        raise RuntimeError(f"Não consegui abrir o produto do SKU {sku} no WooCommerce.")

    frame = admin_frame(page)

    set_featured_image(page, frame, files["variante"])
    add_gallery_images(page, frame, files["fotos"])

    labels_adicionados = []

    if files["datasheet"]:
        add_download_file(page, frame, "Datasheet", files["datasheet"])
        labels_adicionados.append("Datasheet")
    else:
        print("  ! Aviso: datasheet não encontrado na pasta.")

    if files["desenho3d"]:
        add_download_file(page, frame, "Desenho 3D", files["desenho3d"])
        labels_adicionados.append("Desenho 3D")
    else:
        print("  Desenho 3D não encontrado na pasta (ok, é opcional) - pulando.")

    save_product(
        page, frame,
        esperado_imagem=bool(files["variante"]),
        esperado_fotos=len(files["fotos"]),
        esperado_labels=labels_adicionados,
    )
    print(f"  OK - SKU {sku} atualizado com sucesso.\n")


# --------------------------- planilha de SKUs ---------------------------

_COLUNAS_SKU_CONHECIDAS = ("sku", "código", "codigo", "cod", "cod.", "código cigam", "codigo cigam")


def _normalizar_valor_sku(val) -> str:
    """Evita que um SKU numérico vire '117009000372.0' por causa do Excel."""
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()


def load_skus_from_file(path: Path) -> list[str]:
    """Lê os SKUs de uma planilha .xlsx/.xlsm ou .csv. Procura uma coluna
    chamada SKU/Código; se não achar, usa a primeira coluna."""
    ext = path.suffix.lower()
    skus: list[str] = []

    if ext in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook

        wb = load_workbook(path, data_only=True)
        ws = wb.active
        rows_iter = ws.iter_rows(values_only=True)
        header = next(rows_iter, None)
        if header is None:
            return []
        header_lower = [str(h).strip().lower() if h is not None else "" for h in header]

        col_idx = next((i for i, h in enumerate(header_lower) if h in _COLUNAS_SKU_CONHECIDAS), 0)
        print(f"  (lendo SKUs da coluna '{header[col_idx]}')")

        for row in rows_iter:
            if col_idx < len(row) and row[col_idx] not in (None, ""):
                skus.append(_normalizar_valor_sku(row[col_idx]))

    elif ext == ".csv":
        import csv

        with open(path, encoding="utf-8-sig", newline="") as f:
            rows = list(csv.reader(f))
        if not rows:
            return []
        header_lower = [h.strip().lower() for h in rows[0]]
        col_idx = next((i for i, h in enumerate(header_lower) if h in _COLUNAS_SKU_CONHECIDAS), 0)
        print(f"  (lendo SKUs da coluna '{rows[0][col_idx]}')")

        for row in rows[1:]:
            if col_idx < len(row) and row[col_idx].strip():
                skus.append(_normalizar_valor_sku(row[col_idx]))

    else:
        raise ValueError(f"Formato de arquivo não suportado: '{ext}' (use .xlsx ou .csv)")

    # remove duplicados mantendo a ordem original
    vistos = set()
    unicos = []
    for s in skus:
        if s not in vistos:
            vistos.add(s)
            unicos.append(s)
    return unicos


def print_and_save_summary(resultados: list[tuple[str, str, str]]):
    ok = [r for r in resultados if r[1] == "OK"]
    erro = [r for r in resultados if r[1] == "ERRO"]

    print("\n" + "=" * 60)
    print("RESUMO DO LOTE")
    print("=" * 60)
    print(f"Total: {len(resultados)}  |  OK: {len(ok)}  |  Com erro: {len(erro)}")
    if erro:
        print("\nSKUs que deram erro:")
        for sku, _, detalhe in erro:
            print(f"  - {sku}: {detalhe}")

    ts = time.strftime("%Y%m%d-%H%M%S")
    log_path = Path(f"./resultado_lote_{ts}.txt")
    try:
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(f"Resultado do lote - {ts}\n")
            f.write(f"Total: {len(resultados)} | OK: {len(ok)} | Com erro: {len(erro)}\n\n")
            for sku, status, detalhe in resultados:
                f.write(f"{sku}\t{status}\t{detalhe}\n")
        print(f"\nRelatório completo salvo em: {log_path}")
    except Exception:
        pass


def run_batch(page: Page, skus: list[str], root: Path, base_url: str):
    resultados: list[tuple[str, str, str]] = []
    total = len(skus)
    for i, sku in enumerate(skus, start=1):
        print(f"\n[{i}/{total}] " + "=" * 50)
        try:
            process_sku(page, sku, root, base_url)
            resultados.append((sku, "OK", ""))
        except Exception as e:
            print(f"  !! ERRO ao processar SKU {sku}: {e}")
            shot_path = save_debug_screenshot(page, sku)
            detalhe = str(e)
            if shot_path:
                print(f"  -> Print da tela salvo em: {shot_path}")
                detalhe += f" (print: {shot_path})"
            resultados.append((sku, "ERRO", detalhe))

    print_and_save_summary(resultados)


def run_interactive(page: Page, root: Path, base_url: str):
    print("Digite o SKU do produto e aperte ENTER. Digite 'sair' para terminar.\n")
    while True:
        sku = input("SKU: ").strip()
        if not sku or sku.lower() in ("sair", "fim", "exit", "quit"):
            break
        try:
            process_sku(page, sku, root, base_url)
        except Exception as e:
            print(f"  !! ERRO ao processar SKU {sku}: {e}")
            shot_path = save_debug_screenshot(page, sku)
            if shot_path:
                print(f"  -> Print da tela salvo em: {shot_path} (me manda esse arquivo se for pedir ajuda)")
            print("  -> Pulando para o próximo SKU.\n")


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Automação de cadastro de produtos CHOICE")
    parser.add_argument("--base-url", default=BASE_URL)
    parser.add_argument("--root", default=ROOT_FOLDER)
    parser.add_argument("--headless", action="store_true", help="roda sem mostrar o navegador")
    args = parser.parse_args()

    root = Path(args.root)

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            headless=args.headless,
            viewport={"width": 1440, "height": 900},
        )
        page = context.pages[0] if context.pages else context.new_page()

        ensure_logged_in(page, args.base_url)

        print("\n=== Automação de Cadastro de Produtos - CHOICE ===")
        print("1 - Processar uma planilha (.xlsx ou .csv) com uma lista de SKUs")
        print("2 - Digitar os SKUs um por um")
        escolha = input("Opção (1/2): ").strip()

        if escolha == "1":
            caminho = input("Caminho da planilha (.xlsx ou .csv): ").strip().strip('"')
            skus = load_skus_from_file(Path(caminho))
            if not skus:
                print("Não encontrei nenhum SKU nessa planilha.")
            else:
                print(f"{len(skus)} SKU(s) encontrados. Começando o lote...\n")
                run_batch(page, skus, root, args.base_url)
        else:
            run_interactive(page, root, args.base_url)

        context.close()
        print("Encerrado.") 


if __name__ == "__main__":
    main()

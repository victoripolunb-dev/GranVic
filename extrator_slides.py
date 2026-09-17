import os
import re
import sys
import json
import pdfplumber

try:
    import pymupdf  # PyMuPDF >= 1.24
except ImportError:
    try:
        import fitz as pymupdf
    except ImportError:
        pymupdf = None

MIN_LADO_PT = 40.0
MIN_AREA_PT = 90.0 * 60.0
MAX_FRACAO_PAGINA = 0.6
MARGEM_LEGENDA_PT = 40.0


def _legenda_imagem(rect, palavras):
    """Coleta o texto do slide que está dentro ou logo abaixo da imagem (legenda literal)."""
    x0, y0, x1, y1 = rect
    colhidas = []
    for w in palavras:
        wx0, wy0, wx1, wy1, texto = w[0], w[1], w[2], w[3], w[4]
        dentro = wx1 > x0 and wx0 < x1 and wy1 > y0 and wy0 < y1
        abaixo = y0 <= wy0 <= y1 + MARGEM_LEGENDA_PT and wx1 > x0 and wx0 < x1
        if dentro or abaixo:
            colhidas.append((round(wy0, 1), round(wx0, 1), texto))
    colhidas.sort()
    texto = " ".join(t for _, _, t in colhidas)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto[:300] or None


def _extrair_imagens(doc, pagina, numero_pagina, pasta_saida, prefixo):
    if pymupdf is None:
        return []
    import hashlib
    page = doc[numero_pagina - 1]
    area_pagina = page.rect.width * page.rect.height
    palavras = page.get_text("words")
    vistas = set()
    imagens = []
    for xref, *_ in page.get_images(full=True):
        if xref in vistas:
            continue
        vistas.add(xref)
        try:
            rects = page.get_image_rects(xref)
        except Exception:
            rects = []
        if not rects:
            continue
        try:
            info = doc.extract_image(xref)
        except Exception:
            continue
        dados, ext = info.get("image"), (info.get("ext") or "png")
        if not dados:
            continue
        for r in rects:
            largura, altura = r.width, r.height
            if largura < MIN_LADO_PT or altura < MIN_LADO_PT or largura * altura < MIN_AREA_PT:
                continue
            if area_pagina and (largura * altura) > area_pagina * MAX_FRACAO_PAGINA:
                # imagem cobre quase a página inteira → é fundo/background do slide, não um diagrama
                continue
            digest = hashlib.md5(dados).hexdigest()[:8]
            nome_arq = "%s_p%02d_%s.%s" % (prefixo, numero_pagina, digest, ext)
            caminho = os.path.join(pasta_saida, nome_arq)
            if not os.path.exists(caminho):
                with open(caminho, "wb") as f:
                    f.write(dados)
            imagens.append({
                "pagina": numero_pagina,
                "arquivo": os.path.abspath(caminho),
                "largura_pt": round(largura, 1),
                "altura_pt": round(altura, 1),
                "bbox": [round(v, 1) for v in (r.x0, r.y0, r.x1, r.y1)],
                "legenda": _legenda_imagem((r.x0, r.y0, r.x1, r.y1), palavras),
            })
    return imagens


def extrair_slides(pdf_path, output_json_path, pasta_imagens=None):
    print(f"[ExtratorSlides] Lendo PDF: {pdf_path}")
    if pasta_imagens is None:
        pasta_imagens = os.path.join(os.path.dirname(os.path.abspath(pdf_path)), "_imagens")
    os.makedirs(pasta_imagens, exist_ok=True)
    prefixo = re.sub(r"[^A-Za-z0-9]+", "_", os.path.splitext(os.path.basename(pdf_path))[0]).strip("_")[:40] or "slide"

    if pymupdf is None:
        print("[ExtratorSlides] AVISO: PyMuPDF não instalado (pip install pymupdf) — imagens/diagramas não serão extraídos.")

    doc = None
    if pymupdf is not None:
        try:
            doc = pymupdf.open(pdf_path)
        except Exception as erro:
            print(f"[ExtratorSlides] AVISO: falha ao abrir com PyMuPDF ({erro}); seguindo só com texto.")

    paginas = []
    total_imagens = 0
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            texto = page.extract_text() or ""
            imagens = []
            if doc is not None:
                try:
                    imagens = _extrair_imagens(doc, page, i + 1, pasta_imagens, prefixo)
                except Exception as erro:
                    print(f"[ExtratorSlides] AVISO: falha ao extrair imagens da página {i + 1}: {erro}")
            total_imagens += len(imagens)
            paginas.append({
                "pagina": i + 1,
                "texto": texto.strip(),
                "imagens": imagens,
            })

    if doc is not None:
        doc.close()

    with open(output_json_path, "w", encoding="utf-8") as f:
        json.dump(paginas, f, ensure_ascii=False, indent=2)

    print(f"[ExtratorSlides] {len(paginas)} páginas extraídas, {total_imagens} imagens/diagramas. Salvo em {output_json_path}")
    return paginas

if __name__ == "__main__":
    if len(sys.argv) > 2:
        extrair_slides(sys.argv[1], sys.argv[2])
    else:
        print("Uso: python extrator_slides.py <slide.pdf> <saida.json>")
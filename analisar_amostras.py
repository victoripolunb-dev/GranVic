import argparse
import json
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

from docx_fingerprint import comparar_fingerprints, extrair_fingerprint, sao_mesma_aula, alinhar_para_aula_diferente

RAIZ = Path(__file__).resolve().parent.parent
PADRAO_MODELOS = "Modelos Gran Cursos"


def _pasta_modelos_default():
    """Pasta 'Modelos Gran Cursos': dentro do repositório (fluxo GitHub) ou path local legado."""
    candidatos = [
        Path(__file__).resolve().parent / PADRAO_MODELOS,
        Path(r"C:\Users\Victor\Desktop\Repositório - GranVic\Modelos Gran Cursos"),
    ]
    for c in candidatos:
        if c.is_dir():
            return str(c)
    return str(candidatos[0])

CAMPOS_CATEGORIA = (
    "fonte",
    "tamanho_pt",
    "negrito",
    "italico",
    "sublinhado",
    "maiusculas",
    "cor",
    "alinhamento",
    "espaco_antes_pt",
    "espaco_depois_pt",
    "entrelinha",
    "entrelinha_regra",
    "recuo_esquerda_pt",
    "recuo_primeira_linha_pt",
)


def modo(valores):
    vals = [v for v in valores if v is not None]
    if not vals:
        return None
    return Counter(vals).most_common(1)[0][0]


def descobrir_amostras(raiz, pasta_modelos):
    amostras = []
    candidatas = list(raiz.glob("Aula*/")) + list(pasta_modelos.glob("Aula*/")) if pasta_modelos and pasta_modelos.is_dir() else list(raiz.glob("Aula*/"))
    for pasta in candidatas:
        if not pasta.is_dir():
            continue
        for arq in sorted(pasta.glob("*.docx")):
            if arq.name.startswith("~$"):
                continue
            amostras.append({"pasta": pasta.name, "arquivo": arq.name, "caminho": arq})
    return amostras


def consolidar_categoria(fingerprints, chave):
    campos = {}
    for campo in CAMPOS_CATEGORIA:
        colhidos = []
        for fp in fingerprints:
            valor = fp.get("categorias_paragrafo", {}).get(chave, {}).get(campo)
            if valor is not None:
                colhidos.append(valor)
        campos[campo] = modo(colhidos)
    return campos


def consolidar_caixas(fingerprints):
    rotulos = set()
    for fp in fingerprints:
        rotulos |= set(fp.get("caixas_destaque", {}).keys())
    resultado = {}
    for rotulo in sorted(rotulos):
        if rotulo == "sem_rotulo":
            continue
        quadro = {
            "fundo_cor": [],
            "borda_esquerda_estilo": [],
            "borda_esquerda_cor": [],
            "borda_esquerda_largura_pt": [],
            "rotulo_fonte": [],
            "rotulo_tamanho_pt": [],
            "rotulo_negrito": [],
            "rotulo_cor": [],
            "rotulo_maiusculas": [],
            "corpo_fonte": [],
            "corpo_tamanho_pt": [],
            "corpo_negrito": [],
            "corpo_italico": [],
            "alinhamento_corpo": [],
            "quantidade": [],
        }
        for fp in fingerprints:
            cx = fp.get("caixas_destaque", {}).get(rotulo)
            if not cx:
                continue
            borda_esq = (cx.get("bordas") or {}).get("left") or {}
            quadro["fundo_cor"].append(cx.get("fundo_cor"))
            quadro["borda_esquerda_estilo"].append(borda_esq.get("estilo"))
            quadro["borda_esquerda_cor"].append(borda_esq.get("cor"))
            quadro["borda_esquerda_largura_pt"].append(borda_esq.get("largura_pt"))
            fr = cx.get("formato_rotulo") or {}
            fc = cx.get("formato_corpo") or {}
            quadro["rotulo_fonte"].append(fr.get("fonte"))
            quadro["rotulo_tamanho_pt"].append(fr.get("tamanho_pt"))
            quadro["rotulo_negrito"].append(fr.get("negrito"))
            quadro["rotulo_cor"].append(fr.get("cor"))
            quadro["rotulo_maiusculas"].append(fr.get("maiusculas"))
            quadro["corpo_fonte"].append(fc.get("fonte"))
            quadro["corpo_tamanho_pt"].append(fc.get("tamanho_pt"))
            quadro["corpo_negrito"].append(fc.get("negrito"))
            quadro["corpo_italico"].append(fc.get("italico"))
            quadro["alinhamento_corpo"].append(cx.get("alinhamento"))
            quadro["quantidade"].append(cx.get("quantidade"))
        resultado[rotulo] = {
            "fundo_cor": modo(quadro["fundo_cor"]),
            "borda_esquerda_estilo": modo(quadro["borda_esquerda_estilo"]),
            "borda_esquerda_cor": modo(quadro["borda_esquerda_cor"]),
            "borda_esquerda_largura_pt": modo(quadro["borda_esquerda_largura_pt"]),
            "rotulo_formato": {
                "fonte": modo(quadro["rotulo_fonte"]),
                "tamanho_pt": modo(quadro["rotulo_tamanho_pt"]),
                "negrito": modo(quadro["rotulo_negrito"]),
                "cor": modo(quadro["rotulo_cor"]),
                "maiusculas": modo(quadro["rotulo_maiusculas"]),
            },
            "corpo_formato": {
                "fonte": modo(quadro["corpo_fonte"]),
                "tamanho_pt": modo(quadro["corpo_tamanho_pt"]),
                "negrito": modo(quadro["corpo_negrito"]),
                "italico": modo(quadro["corpo_italico"]),
            },
            "alinhamento_corpo": modo(quadro["alinhamento_corpo"]),
            "quantidade_media": round(sum(q for q in quadro["quantidade"] if q) / max(len([q for q in quadro["quantidade"] if q]), 1), 1)
            if any(quadro["quantidade"])
            else None,
        }
    return resultado


def consolidar_icones(fingerprints):
    chaves = set()
    for fp in fingerprints:
        chaves |= set(fp.get("icones", {}).keys())
    campos = (
        "fonte", "tamanho_pt", "negrito", "italico", "sublinhado", "maiusculas",
        "cor", "alinhamento", "espaco_antes_pt", "espaco_depois_pt", "entrelinha", "entrelinha_regra",
    )
    mapa = {}
    for chave in sorted(chaves):
        colhidos = {c: [] for c in campos}
        quant = 0
        for fp in fingerprints:
            dado = fp.get("icones", {}).get(chave)
            if not dado:
                continue
            quant += dado.get("quantidade") or 0
            for c in campos:
                if dado.get(c) is not None:
                    colhidos[c].append(dado[c])
        mapa[chave] = {"quantidade_total": quant}
        mapa[chave].update({c: modo(colhidos[c]) for c in campos})
    return mapa


def consolidar_cores(fingerprints):
    coletor = Counter()
    usos = {}
    for fp in fingerprints:
        for chave in ("titulo_1", "titulo_2", "titulo_3", "corpo"):
            cor = fp.get(chave, {}).get("cor")
            if cor:
                coletor[cor] += 1
                usos.setdefault(cor, set()).add(chave)
        for rotulo, cx in fp.get("caixas_destaque", {}).items():
            if cx.get("fundo_cor"):
                coletor[cx["fundo_cor"]] += 1
                usos.setdefault(cx["fundo_cor"], set()).add("caixa:%s" % rotulo)
    lista = [
        {"hex": cor, "frequencia": n, "usos": sorted(usos.get(cor, []))}
        for cor, n in coletor.most_common()
    ]
    return lista


def consolidar_secoes(fingerprints):
    margens = {lado: [] for lado in ("superior", "inferior", "esquerda", "direita")}
    larguras = []
    orientacao = []
    for fp in fingerprints:
        sec1 = (fp.get("secoes") or [{}])[0]
        m = sec1.get("margens_pt") or {}
        for lado in margens:
            if m.get(lado) is not None:
                margens[lado].append(m[lado])
        if sec1.get("largura_pt"):
            larguras.append(sec1["largura_pt"])
        if sec1.get("orientacao"):
            orientacao.append(sec1["orientacao"])
    return {
        "margens_pt": {lado: modo(lista) for lado, lista in margens.items()},
        "largura_pt": modo(larguras),
        "orientacao": modo(orientacao),
    }


def construir_regras(amostras):
    fingerprints = []
    detalhe = []
    avisos = []
    for amostra in amostras:
        caminho = amostra["caminho"]
        try:
            fp = extrair_fingerprint(caminho)
        except Exception as erro:
            avisos.append("Falha ao analisar %s: %s" % (caminho.name, erro))
            continue
        fingerprints.append(fp)
        detalhe.append(
            {
                "pasta": amostra["pasta"],
                "arquivo": amostra["arquivo"],
                "fingerprint": fp,
            }
        )
    if not fingerprints:
        return None, detalhe, avisos

    estilos = {
        chave: consolidar_categoria(fingerprints, chave)
        for chave in ("titulo_principal", "titulo_1", "titulo_2", "titulo_3", "enfase", "corpo")
    }
    regras = {
        "gerado_em": datetime.now().isoformat(timespec="seconds"),
        "amostras_analisadas": [{"pasta": a["pasta"], "arquivo": a["arquivo"]} for a in amostras],
        "aviso": avisos,
        "template_recomendado": str(amostras[0]["caminho"]),
        "estilos": estilos,
        "cores_corporativas": consolidar_cores(fingerprints),
        "caixas_destaque": consolidar_caixas(fingerprints),
        "icones": consolidar_icones(fingerprints),
        "linha_limitadora": {
            "estilo": modo([fp.get("linha_limitadora", {}).get("estilo") for fp in fingerprints]),
            "cor": modo([fp.get("linha_limitadora", {}).get("cor") for fp in fingerprints]),
            "largura_pt": modo([fp.get("linha_limitadora", {}).get("largura_pt") for fp in fingerprints]),
            "quantidade_total": sum(fp.get("linha_limitadora", {}).get("quantidade", 0) for fp in fingerprints),
        },
        "marcadores_tempo": {
            "quantidade_total": sum(fp.get("marcadores_tempo", {}).get("quantidade", 0) for fp in fingerprints),
        },
        "citacoes_lei": {
            "quantidade_total": sum(fp.get("citacoes_lei", {}).get("quantidade", 0) for fp in fingerprints),
            "fonte": modo([fp.get("citacoes_lei", {}).get("fonte") for fp in fingerprints]),
            "tamanho_pt": modo([fp.get("citacoes_lei", {}).get("tamanho_pt") for fp in fingerprints]),
            "italico": modo([fp.get("citacoes_lei", {}).get("italico") for fp in fingerprints]),
            "cor": modo([fp.get("citacoes_lei", {}).get("cor") for fp in fingerprints]),
            "recuo_esquerda_pt": modo([fp.get("citacoes_lei", {}).get("recuo_esquerda_pt") for fp in fingerprints]),
        },
        "tabelas": {
            "quantidade_media": round(
                sum(fp.get("tabelas", {}).get("quantidade", 0) for fp in fingerprints) / len(fingerprints), 2
            ),
            "exemplo_header": modo([str(fp.get("tabelas", {}).get("exemplos") or "") for fp in fingerprints]),
        },
        "secoes_pagina": consolidar_secoes(fingerprints),
        "desenhos_imagens": {
            "quantidade_media": round(sum(fp.get("desenhos", {}).get("pagina", 0) for fp in fingerprints) / len(fingerprints), 2)
            if fingerprints
            else 0
        },
        "fingerprints": fingerprints,
    }
    return regras, detalhe, avisos


def imprimir_resumo(regras):
    if regras is None:
        print("Nenhuma amostra .docx encontrada.")
        return
    print("Amostras analisadas: %d" % len(regras["amostras_analisadas"]))
    print("Template recomendado: %s" % Path(regras["template_recomendado"]).name)
    print("\n--- Estilos por categoria ---")
    for chave, estilo in regras["estilos"].items():
        print(
            "%-8s  fonte=%s  %spt  negrito=%s  italico=%s  alinh=%s  cor=%s"
            % (
                chave,
                estilo.get("fonte"),
                estilo.get("tamanho_pt"),
                estilo.get("negrito"),
                estilo.get("italico"),
                estilo.get("alinhamento"),
                estilo.get("cor"),
            )
        )
        print(
            "          espaco antes/depois: %spt / %spt   entrelinha: %s (%s)   recuo prim: %spt"
            % (
                estilo.get("espaco_antes_pt"),
                estilo.get("espaco_depois_pt"),
                estilo.get("entrelinha"),
                estilo.get("entrelinha_regra"),
                estilo.get("recuo_primeira_linha_pt"),
            )
        )
    print("\n--- Cores corporativas ---")
    for c in regras["cores_corporativas"][:10]:
        print("  %s  freq=%d  usos=%s" % (c["hex"], c["frequencia"], ",".join(c["usos"])))
    print("\n--- Caixas de destaque ---")
    for rotulo, cx in regras["caixas_destaque"].items():
        print(
            "  %-14s fundo=%s  bordaEsq=%s/%s/%s  rotulo(fonte=%s, %spt, negrito=%s, cor=%s)  corpo(alinh=%s, it=%s)"
            % (
                rotulo,
                cx.get("fundo_cor"),
                cx.get("borda_esquerda_estilo"),
                cx.get("borda_esquerda_cor"),
                cx.get("borda_esquerda_largura_pt"),
                cx["rotulo_formato"].get("fonte"),
                cx["rotulo_formato"].get("tamanho_pt"),
                cx["rotulo_formato"].get("negrito"),
                cx["rotulo_formato"].get("cor"),
                cx.get("alinhamento_corpo"),
                cx["corpo_formato"].get("italico"),
            )
        )
    print("\n--- Ícones detectados ---")
    for icone, ic in regras.get("icones", {}).items():
        print(
            "  %-22s x%d  fonte=%s  %spt  negrito=%s  cor=%s  alinh=%s"
            % (
                icone,
                ic.get("quantidade_total", 0),
                ic.get("fonte"),
                ic.get("tamanho_pt"),
                ic.get("negrito"),
                ic.get("cor"),
                ic.get("alinhamento"),
            )
        )
    lr = regras.get("linha_limitadora", {})
    print("\n--- Linha limitadora (ícones) ---")
    print("  total=%s  estilo=%s  cor=%s  largura=%spt" % (lr.get("quantidade_total"), lr.get("estilo"), lr.get("cor"), lr.get("largura_pt")))
    print("\n--- Marcadores de tempo (relógio) ---")
    print("  total=%s" % regras.get("marcadores_tempo", {}).get("quantidade_total"))
    print("\n--- Citações de lei ---")
    c = regras["citacoes_lei"]
    print("  total=%s  fonte=%s  %spt  italico=%s  recuo=%spt" % (c.get("quantidade_total"), c.get("fonte"), c.get("tamanho_pt"), c.get("italico"), c.get("recuo_esquerda_pt")))
    print("\n--- Página ---")
    s = regras["secoes_pagina"]
    print("  margens=%s  orientacao=%s" % (s["margens_pt"], s.get("orientacao")))


def rodar_comparar(gerado, amostra, force_modo=None):
    fp_gerado = extrair_fingerprint(gerado)
    fp_amostra = extrair_fingerprint(amostra)
    mesma_aula = force_modo is None and sao_mesma_aula(fp_amostra, fp_gerado)
    modo = force_modo or ("mesma_aula" if mesma_aula else "aula_diferente")
    if modo == "aula_diferente":
        fp_amostra, fp_gerado = alinhar_para_aula_diferente(fp_amostra, fp_gerado)
    linhas = comparar_fingerprints(fp_amostra, fp_gerado, modo=modo)
    divergencias = [l for l in linhas if l["status"] == "DIVERGENTE"]
    ignoradas = [l for l in linhas if l.get("ignorada")]
    ok = [l for l in linhas if l["status"] == "OK"]
    print("Comparação: %s  vs  %s  (modo=%s)" % (Path(amostra).name, Path(gerado).name, modo))
    print("Propriedades auditadas: %d  |  OK: %d  |  Divergentes: %d  |  Ignoradas (aula diferente): %d"
          % (len(linhas), len(ok), len(divergencias), len(ignoradas)))
    for l in divergencias:
        print("  [%s] %s  |  amostra=%s  gerado=%s" % (l["status"], l["propriedade"], l["amostra"], l["gerado"]))
    if divergencias:
        print("\nAUDITORIA: FALHOU (%d divergências)" % len(divergencias))
    else:
        print("\nAUDITORIA: OK — 100%% de fidelidade nas propriedades auditadas")
    return linhas


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    parser = argparse.ArgumentParser(description="Analisa amostras .docx do Gran Cursos e mapeia o padrão de estilo.")
    parser.add_argument("--raiz", default=str(RAIZ), help="Pasta raiz das entregas (padrão: Entregas Gran Cursos)")
    parser.add_argument("--modelos", default=_pasta_modelos_default(), help="Pasta com as amostras-modelo (padrão: Modelos Gran Cursos do repositório)")
    parser.add_argument("--saida", default=str(RAIZ / "_sistema" / "regras_estilo.json"), help="Caminho do regras_estilo.json")
    parser.add_argument("--detalhe", default=str(RAIZ / "_sistema" / "inspecao_detalhada.json"), help="Caminho do relatório detalhado")
    parser.add_argument("--comparar", metavar="GERADO", help="Audita um material gerado comparando com a amostra")
    parser.add_argument("--amostra", metavar="AMOSTRA", help="Amostra a usar na auditoria (padrão: template_recomendado)")
    parser.add_argument("--modo", choices=["mesma_aula", "aula_diferente"], help="Força o modo de comparação (padrão: automático)")
    args = parser.parse_args()

    raiz = Path(args.raiz)
    modelos = Path(args.modelos)

    if args.comparar:
        gerado = Path(args.comparar)
        amostra = Path(args.amostra) if args.amostra else None
        if amostra is None:
            saida_regras = Path(args.saida)
            if saida_regras.exists():
                dados = json.loads(saida_regras.read_text(encoding="utf-8"))
                amostra = Path(dados.get("template_recomendado", ""))
            if amostra is None or not amostra.is_file():
                print("Informe --amostra ou rode primeiro a análise para ter o template_recomendado.")
                sys.exit(2)
        linhas = rodar_comparar(gerado, amostra, force_modo=args.modo)
        auditoria = {
            "gerado_em": datetime.now().isoformat(timespec="seconds"),
            "amostra": str(amostra),
            "gerado": str(gerado),
            "modo": args.modo or ("mesma_aula" if sao_mesma_aula(extrair_fingerprint(str(amostra)), extrair_fingerprint(str(gerado))) else "aula_diferente"),
            "resultado": "OK" if not [l for l in linhas if l["status"] == "DIVERGENTE"] else "DIVERGENTE",
            "propriedades": linhas,
        }
        caminho_auditoria = raiz / "_sistema" / "auditoria_regras.json"
        caminho_auditoria.write_text(json.dumps(auditoria, ensure_ascii=False, indent=2), encoding="utf-8")
        print("Relatório de auditoria salvo em %s" % caminho_auditoria)
        return

    amostras = descobrir_amostras(raiz, modelos)
    if not amostras:
        print(
            "Nenhum arquivo .docx encontrado em %s/Aula* nem em %s/Aula*.\n"
            "Coloque as amostras (ex.: 'Aula 1 - Documento pronto ...docx') e rode novamente."
            % (raiz, modelos)
        )
        sys.exit(1)

    regras, detalhe, avisos = construir_regras(amostras)
    caminho_saida = Path(args.saida)
    caminho_saida.parent.mkdir(parents=True, exist_ok=True)
    if regras is not None:
        caminho_saida.write_text(json.dumps(regras, ensure_ascii=False, indent=2), encoding="utf-8")
        print("Regras salvas em %s" % caminho_saida)
    caminho_detalhe = Path(args.detalhe)
    caminho_detalhe.parent.mkdir(parents=True, exist_ok=True)
    caminho_detalhe.write_text(
        json.dumps({"gerado_em": datetime.now().isoformat(timespec="seconds"), "amostras": detalhe, "avisos": avisos}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print("Inspeção detalhada em %s" % caminho_detalhe)
    imprimir_resumo(regras)


if __name__ == "__main__":
    main()
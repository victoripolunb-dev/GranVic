import os
import sys
import json
from pathlib import Path
from docx_fingerprint import extrair_fingerprint, comparar_fingerprints, sao_mesma_aula, alinhar_para_aula_diferente

def rodar_auditoria(doc_gerado, doc_referencia, saida_relatorio, mesma_aula=None, modo=None):
    print(f"[Auditoria] Analisando gerado: {doc_gerado}")
    print(f"[Auditoria] Referência: {doc_referencia}")

    fp_gerado = extrair_fingerprint(doc_gerado)
    fp_ref = extrair_fingerprint(doc_referencia)

    if mesma_aula is None:
        mesma_aula = sao_mesma_aula(fp_ref, fp_gerado)
    if modo is None:
        modo = "mesma_aula" if mesma_aula else "aula_diferente"

    if modo == "aula_diferente":
        fp_ref, fp_gerado = alinhar_para_aula_diferente(fp_ref, fp_gerado)
    linhas = comparar_fingerprints(fp_ref, fp_gerado, modo=modo)
    divergencias = [l for l in linhas if l["status"] == "DIVERGENTE"]
    ignoradas = [l for l in linhas if l.get("ignorada")]

    relatorio = {
        "status": "APROVADO" if len(divergencias) == 0 else f"FALHOU ({len(divergencias)} divergências)",
        "mesma_aula": mesma_aula,
        "modo_comparacao": modo,
        "total_divergencias": len(divergencias),
        "divergencias_ignoradas_quantidade_conteudo": len(ignoradas),
        "divergencias": divergencias,
        "ignoradas": ignoradas,
    }

    with open(saida_relatorio, "w", encoding="utf-8") as f:
        json.dump(relatorio, f, ensure_ascii=False, indent=2)

    print(f"[Auditoria] Aulas {'iguais' if mesma_aula else 'DIFERENTES'} (modo={modo})")
    print(f"[Auditoria] Resultado: {relatorio['status']}  (ignoradas por aula diferente: {len(ignoradas)})")
    print(f"[Auditoria] Relatório salvo em {saida_relatorio}")
    return relatorio


def encontrar_amostra_correspondente(doc_gerado, pasta_modelos, extensao="*.docx"):
    """Escolhe, entre as amostras dos modelos, a que corresponde à aula do documento gerado.
    Critérios em ordem: 1) mesma aula por cabeçalho/título; 2) mesma pasta 'Aula N'; 3) 1ª amostra."""
    import re
    pasta_modelos = Path(pasta_modelos)
    amostras = [a for a in sorted(pasta_modelos.rglob(extensao)) if not a.name.startswith("~$")]
    if not amostras:
        return None

    fp_gerado = extrair_fingerprint(doc_gerado)
    for arq in amostras:
        try:
            if sao_mesma_aula(extrair_fingerprint(str(arq)), fp_gerado):
                return str(arq)
        except Exception:
            continue

    m_ger = re.search(r"aula\s*0?([1-9])", str(doc_gerado).lower())
    if m_ger:
        for arq in amostras:
            if re.search(r"aula\s*0?%s" % m_ger.group(1), str(arq.parent).lower()):
                return str(arq)
    return str(amostras[0])


if __name__ == "__main__":
    if len(sys.argv) > 3:
        doc_g = sys.argv[1]
        doc_r = sys.argv[2]
        rel = sys.argv[3]
        modo = None
        if "--modo" in sys.argv:
            modo = sys.argv[sys.argv.index("--modo") + 1]
        rodar_auditoria(doc_g, doc_r, rel, mesma_aula=(modo == "mesma_aula"), modo=modo)
    else:
        print("Uso: python auditoria.py <doc_gerado.docx> <doc_referencia.docx> <relatorio.json> [--modo mesma_aula|aula_diferente]")
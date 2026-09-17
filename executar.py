import os
import sys
import json
import argparse
from pathlib import Path
from transcrever import transcrever_audio
from extrator_slides import extrair_slides
from processador_gemini import processar_conteudo
from gerador_docx import gerar_docx
from auditoria import rodar_auditoria, encontrar_amostra_correspondente

def _raiz_repositorio():
    return Path(__file__).resolve().parent


def _localizar_pasta_modelos():
    """Localiza a pasta 'Modelos Gran Cursos': primeiro ao lado do repositório
    (fluxo do GitHub), depois no path local original de desenvolvimento."""
    candidatos = [
        _raiz_repositorio() / "Modelos Gran Cursos",
        Path(r"C:\Users\Victor\Desktop\Repositório - GranVic\Modelos Gran Cursos"),
    ]
    for c in candidatos:
        if c.is_dir():
            return str(c)
    return str(candidatos[0])


def _resolver_template():
    """Procura o template 'Documento pronto' da Aula 1 dentro da pasta de modelos."""
    base = Path(PASTA_MODELOS)
    nome_desejado = "Aula 1 - Documento pronto - Planejamento, execução, registro, monitoramento e avaliação de ações socioeducativas.docx"
    for sub in sorted(base.glob("Aula 1*")):
        for f in sub.glob("*.docx"):
            if f.name == nome_desejado:
                return str(f)
            if "Documento pronto" in f.name:
                return str(f)
    return str(base / "Aula 1" / nome_desejado)


PASTA_MODELOS = _localizar_pasta_modelos()
TEMPLATE_PADRAO = _resolver_template()

def _normalizar_estruturado(blocos):
    """Converte estrutura legada (sem campo 'tipo') para o schema atual de blocos."""
    tem_tipo = any(isinstance(b, dict) and "tipo" in b for b in blocos)
    if tem_tipo:
        return blocos
    padrao_icone = {
        "alerta": "atencao", "conceito": "dica", "destaque": "resumo", "detalhar": "resumo",
    }
    novos = []
    for b in blocos:
        if not isinstance(b, dict):
            continue
        if b.get("titulo"):
            novos.append({"tipo": "titulo", "texto": b["titulo"]})
        if b.get("questao"):
            novos.append({"tipo": "questao", "texto": b["questao"]})
            alt = b.get("alternativa")
            if alt:
                novos.append({"tipo": "alternativa", "letra": "", "texto": alt})
        if b.get("enfase"):
            novos.append({"tipo": "enfase", "texto": b["enfase"]})
        if b.get("icone"):
            nome = (b["icone"] or "").strip().lower().replace(" ", "_")
            novos.append({"tipo": "icone", "nome": padrao_icone.get(nome, nome)})
        if b.get("corpo"):
            novos.append({"tipo": "corpo", "texto": b["corpo"]})
        if b.get("gabarito_fim"):
            novos.append({"tipo": "gabarito_fim", "respostas": [str(b["gabarito_fim"])]})
    return novos

def pipeline_completo(pasta_aula, amostra=None, limpar_materiais=False):
    pasta_aula = Path(pasta_aula)
    print(f"\n========================================")
    print(f" INICIANDO PIPELINE GRAN CURSOS: {pasta_aula}")
    print(f"========================================")
    
    # 1. Localiza mp3/mp4 e pdf na pasta
    arquivos = os.listdir(pasta_aula)
    audio_file = next((os.path.join(pasta_aula, f) for f in arquivos if f.lower().endswith((".mp3", ".mp4", ".wav", ".m4a"))), None)
    pdf_file = next((os.path.join(pasta_aula, f) for f in arquivos if f.lower().endswith(".pdf")), None)
    
    if not audio_file or not pdf_file:
        print(f"[Erro] A pasta deve conter exatamente 1 áudio (.mp3/.mp4) e 1 PDF (.pdf). Encontrados: Áudio={audio_file}, PDF={pdf_file}")
        return
        
    base_name = os.path.join(pasta_aula, "pipeline")
    trans_json = base_name + "_transcricao.json"
    slides_json = base_name + "_slides.json"
    estruturado_json = base_name + "_estruturado.json"
    saida_docx = os.path.join(pasta_aula, "material_final.docx")
    relatorio_json = os.path.join(pasta_aula, "relatorio_auditoria.json")
    
    # Referência (amostra/template) a usar na geração e na auditoria
    template_path = amostra or TEMPLATE_PADRAO
    if not os.path.exists(template_path):
        print(f"[Erro] Amostra/template não encontrado: {template_path}")
        return
    
    # Passo 1: Transcrever áudio
    if not os.path.exists(trans_json):
        transcrever_audio(audio_file, trans_json)
    else:
        print("[Pipeline] Transcrição já existente, pulando...")
        
    # Passo 2: Extrair slides
    if not os.path.exists(slides_json):
        extrair_slides(pdf_file, slides_json)
    else:
        print("[Pipeline] Slides extraídos já existentes, pulando...")
        
    # Passo 3: Processar com Gemini
    if not os.path.exists(estruturado_json):
        processar_conteudo(trans_json, slides_json, estruturado_json)
    else:
        print("[Pipeline] Conteúdo estruturado já existente, pulando...")
        
    # Passo 4: Gerar Word (.docx)
    with open(estruturado_json, "r", encoding="utf-8") as f:
        blocos_brutos = json.load(f)
    blocos = _normalizar_estruturado(blocos_brutos)
    if blocos is not blocos_brutos and blocos_brutos:
        versao = os.path.join(pasta_aula, "pipeline_estruturado_normalizado.json")
        with open(versao, "w", encoding="utf-8") as f:
            json.dump(blocos, f, ensure_ascii=False, indent=2)
        print(f"[Pipeline] Schema legado detectado, normalizado salvo em: {versao}")
    gerar_docx(template_path, estruturado_json, saida_docx, blocos=blocos)
    
    # Passo 5: Rodar Auditoria (usa a amostra correspondente à aula, se existir)
    amostra_auditoria = amostra or encontrar_amostra_correspondente(saida_docx, PASTA_MODELOS)
    if amostra_auditoria:
        rodar_auditoria(saida_docx, amostra_auditoria, relatorio_json)
    else:
        print(f"[Pipeline] Nenhuma amostra encontrada; auditoria pulada.")
    
    print(f"\n[Sucesso] Pipeline concluído! Documento gerado em:\n  -> {saida_docx}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Pipeline de degravação Gran Cursos")
    parser.add_argument("pasta_aula", help="Caminho da pasta da aula (com áudio e PDF)")
    parser.add_argument("--amostra", help="Amostra/template .docx a usar na geração e auditoria")
    args = parser.parse_args()
    pipeline_completo(args.pasta_aula, amostra=args.amostra)

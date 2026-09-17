import os
import re
import json
import sys
from dotenv import load_dotenv
from google import genai
from google.genai import types

load_dotenv()


def _cliente():
    chave = os.environ.get("GEMINI_API_KEY")
    if not chave:
        raise RuntimeError("GEMINI_API_KEY não configurada no .env")
    return genai.Client(api_key=chave)

PROMPT_SISTEMA = """
Você é um redator sênior especializado no padrão editorial Gran Cursos para degravações educacionais.
Seu objetivo é transformar a transcrição bruta da aula falada e o conteúdo dos slides em um documento estruturado, MUITO EXTENSO, DENSO E ABRANGENTE (com pelo menos 80 a 110 parágrafos e mais de 2500 palavras), cobrindo minuciosamente cada explanação, conceito, raciocínio, exemplo e desdobramento apresentado pelo professor do início ao fim da aula.

REGRAS EDITORIAIS:
1. EXPANSÃO E PROFUNDIDADE MÁXIMA: Degrave de forma exaustiva e extremamente detalhada. Nunca resuma! Para cada tema abordado na fala do professor, redija múltiplos parágrafos de corpo longos, explicativos e robustos. Explore cada conceito a fundo, exatamente como nos materiais oficiais de referência do Gran Cursos. Remova apenas a oralidade pura ("olá pessoal", "né", gaguejamentos).
2. TÍTULO PRINCIPAL: O título da aula deve vir em letras maiúsculas e em negrito.
3. PARÁGRAFOS DE CORPO: Redigidos em texto corrido (não em tópicos bullet points, exceto itens de questões). Cada parágrafo deve ser rico em conteúdo técnico e conceitual.
4. ÍCONES OBRIGATÓRIOS (em negrito isolado no parágrafo antes do conteúdo). Use SOMENTE estes nomes padronizados (sem duas palavras em "Comentário:" dentro do texto — o rótulo é um bloco separado):
   - "comentario" (explicações do professor sobre questões ou conceitos)
   - "direto_do_concurso" (antes de enunciados de questões)
   - "exercicios_de_fixacao", "o_pulo_do_gato", "atencao", "pegadinha_da_banca", "relembrando", "resolucao", "obs"
   - "gabarito" (somente no final, com listagem)
5. LINHA LIMITADORA: Após cada bloco de ícone/comentário, o sistema insere uma linha divisória horizontal. Você apenas marca onde começa e termina o bloco.
6. MARCADORES DE TEMPO: Inserir a cada aproximadamente 5 minutos no formato " 5MIN", " 10MIN", etc., no final do parágrafo correspondente.
7. CITAÇÕES DE LEI: use o bloco {"tipo":"citacao_lei"} para artigos/parágrafos de lei que o professor citar (o sistema aplica recuo, fonte menor e negrito nos artigos).
8. GABARITO FINAL: gabarito com todas as respostas estritamente no final, letra por letra (letra minúscula; C/E maiúsculos para certo/errado).
9. EXATAS (matemática, estatística, probabilidade, física, contabilidade):
   - Fórmulas e expressões matemáticas SEMPRE em bloco {"tipo":"formula"} (o sistema centraliza em parágrafo próprio). Use símbolos Unicode legíveis (×, ÷, √, ∩, ∪, ≤, ≥, ≠, Σ, μ, σ) e sublinhe potências/expoentes com ^ (ex.: "x^2", "10^-3") — o sistema converte.
   - Dados em formato tabular (comparações, séries, amostras, propriedades) em bloco {"tipo":"tabela"} com {"cabecalho":["..."],"linhas":[["...","..."]]}. Só crie tabela se a informação realmente se beneficia de colunas.
   - Testes/questões com enunciado e alternativas: use "questao"/"alternativa" normalmente.
10. DIAGRAMAS/IMAGENS: Se o slide contiver diagrama/figura que o professor menciona ou explica, emita {"tipo":"imagem","pagina":N,"imagem":M,"legenda":"descrição literal do slide ou do professor"}.
11. NÚMEROS DE QUESTÃO: numerar apenas se o professor numerar; nomenclatura BANCA/CARGO/ÓRGÃO/ANO em caixa-alta separada por /, sem espaços, na ordem do professor.
12. Não use markdown (nem **, nem #, nem __  ) em nenhum texto.

FORMATO DE SAÍDA:
Retorne EXATAMENTE um JSON válido (sem markdown extra fora do json) contendo uma lista de blocos:
[
  {"tipo":"titulo","texto":"TEXTO DO TÍTULO"},
  {"tipo":"corpo","texto":"Parágrafo muito longo e detalhado..."},
  {"tipo":"enfase","texto":"Texto em destaque."},
  {"tipo":"icone","nome":"direto_do_concurso"},
  {"tipo":"questao","texto":"Enunciado..."},
  {"tipo":"alternativa","letra":"a","texto":"Texto da alternativa..."},
  {"tipo":"icone","nome":"comentario"},
  {"tipo":"comentario","texto":"Comentário extenso e aprofundado..."},
  {"tipo":"citacao_lei","texto":"Art. 25. Verificada a infração..."},
  {"tipo":"formula","texto":"P(X ∩ Y) = P(X) × P(Y)"},
  {"tipo":"tabela","cabecalho":["Média","Erro"],"linhas":[["10","0,5"],["12","0,6"]]},
  {"tipo":"imagem","pagina":1,"imagem":1,"legenda":"Fluxograma do processo"},
  {"tipo":"tempo","minuto":5},
  {"tipo":"gabarito_fim","respostas":["a","b","e"]}
]
"""

PADRAO_MARKDOWN = re.compile(r"(\*\*|__|\*+|_)(.+?)\1")
ROTULOS_DUPLICADOS = (
    "comentario", "direto_do_concurso", "exercicios_de_fixacao", "o_pulo_do_gato",
    "atencao", "pegadinha_da_banca", "relembrando", "resolucao", "obs", "gabarito",
)
SUBST_ACENTOS = {
    ord("á"): "a", ord("é"): "e", ord("í"): "i", ord("ó"): "o", ord("ú"): "u",
    ord("à"): "a", ord("ã"): "a", ord("õ"): "o", ord("â"): "a", ord("ê"): "e",
    ord("ô"): "o", ord("Á"): "A", ord("É"): "E", ord("Í"): "I", ord("Ó"): "O",
    ord("Ú"): "U", ord("À"): "A", ord("Ã"): "A", ord("Õ"): "O", ord("Â"): "A",
    ord("Ê"): "E", ord("Ô"): "O", ord("Ç"): "C", ord("ç"): "c",
}
MAP_ICONE = {
    "comentario": "comentario", "comment": "comentario", "comentário": "comentario",
    "direto do concurso": "direto_do_concurso", "direto-do-concurso": "direto_do_concurso",
    "exercicios de fixacao": "exercicios_de_fixacao", "exercícios de fixação": "exercicios_de_fixacao",
    "o pulo do gato": "o_pulo_do_gato", "pulo do gato": "o_pulo_do_gato",
    "atencao": "atencao", "atenção": "atencao", "atenção!": "atencao",
    "pegadinha da banca": "pegadinha_da_banca", "pegadinha": "pegadinha_da_banca",
    "relembrando": "relembrando", "resolucao": "resolucao", "resolução": "resolucao",
    "obs": "obs", "obs.": "obs", "gabarito": "gabarito",
    "alerta": "atencao", "aviso": "atencao", "cuidado": "atencao",
    "conceito": "dica", "definicao": "dica", "destaque": "resumo",
    "detalhar": "resumo", "resumo": "resumo",
}


def sanitizar_texto(texto):
    if not texto:
        return texto
    texto = PADRAO_MARKDOWN.sub(r"\2", texto)
    texto = texto.replace("**", "").replace("__", "")
    while texto != (t := texto.replace(" * ", " ")):
        texto = t
    return texto.strip()


def normalizar_nome_icone(nome):
    if not nome:
        return None
    n = nome.translate(SUBST_ACENTOS).strip().lower()
    n = re.sub(r"[^a-z0-9\s]", " ", n).strip()
    n = re.sub(r"\s+", " ", n)
    return MAP_ICONE.get(n)


def remover_rotulo_duplicado(texto, nome_icone=None):
    t = (texto or "").strip()
    if not t:
        return t
    plano = t.translate(SUBST_ACENTOS).lower()
    for rotulo in ROTULOS_DUPLICADOS:
        rot = re.escape(rotulo.replace("_", " "))
        m_com_pont = re.match(r"\s*%s\s*[:\-–—.\s]*" % rot, plano)
        m_nu = re.match(r"\s*%s\s*$" % rot, plano)
        fim = None
        if m_com_pont:
            fim = m_com_pont.end()
        elif nome_icone and rotulo.replace("_", " ") == nome_icone.replace("_", " ") and m_nu:
            fim = m_nu.end()
        if fim:
            resto = t[fim:].strip()
            if resto:
                return resto
    return t


def sanitizar_blocos(blocos):
    saida = []
    for i, bloco in enumerate(blocos):
        if not isinstance(bloco, dict):
            continue
        bloco = dict(bloco)
        tipo = (bloco.get("tipo") or "").strip().lower()
        if tipo == "ícone" or tipo == "icone":
            tipo = "icone"
        if tipo == "gabarito":
            tipo = "gabarito_fim"
        if tipo == "tempo" and "minuto" not in bloco and "texto" in bloco:
            m = re.search(r"(\d{1,2})\s*min", str(bloco["texto"]), re.IGNORECASE)
            if m:
                bloco["minuto"] = int(m.group(1))
        if tipo == "icone":
            bloco["nome"] = normalizar_nome_icone(bloco.get("nome") or bloco.get("texto"))
            if bloco.get("texto"):
                bloco["texto"] = sanitizar_texto(bloco["texto"])
            saida.append(bloco)
            continue
        if "texto" in bloco:
            bloco["texto"] = sanitizar_texto(bloco["texto"])
        if tipo in ("corpo", "comentario", "questao", "enfase", "citacao_lei", "formula"):
            nome = None
            if i > 0 and isinstance(saida[-1], dict) and saida[-1].get("tipo") == "icone":
                nome = saida[-1].get("nome")
            bloco["texto"] = remover_rotulo_duplicado(bloco["texto"], nome)
        if tipo == "alternativa" and "letra" not in bloco and "texto" in bloco:
            m = re.match(r"^\s*([a-eA-E])[)\s.:–-]\s*", bloco["texto"])
            if m:
                bloco["letra"] = m.group(1).lower()
                bloco["texto"] = bloco["texto"][m.end():].strip()
        saida.append(bloco)
    return saida


def processar_conteudo(transcricao_json_path, slides_json_path, saida_json_path):
    print("[ProcessadorGemini] Lendo dados...")
    with open(transcricao_json_path, "r", encoding="utf-8") as f:
        trans = json.load(f)
    with open(slides_json_path, "r", encoding="utf-8") as f:
        slides = json.load(f)
        
    # Pega a transcrição inteira sem truncar para garantir 100% de fidelidade ao áudio longo
    trans_completa = trans.get('text', '')
    
    prompt_usuario = f"""
Aqui está o conteúdo extraído da aula:

--- SLIDES DA AULA ---
{json.dumps(slides, ensure_ascii=False, indent=2)}

--- TRANSCRIÇÃO DA AULA (ÁUDIO COMPLETO) ---
{trans_completa}

INSTRUÇÃO CRÍTICA DE EXTENSÃO MÁXIMA (PADRÃO DOS SEUS DOCUMENTOS):
Para que o material fique idêntico aos seus documentos de referência (que possuem mais de 100 parágrafos e cerca de 2.700 palavras), você DEVE:
1. Escrever uma quantidade massiva de parágrafos de corpo (pelo menos 80 a 100 parágrafos estruturados).
2. Desdobrar e detalhar minuciosamente cada frase, raciocínio, exemplo e conceito explanado na transcrição do áudio.
3. Não economize palavras nem reduza explicações. Transforme cada trecho da aula em ricas explicações acadêmicas e pedagógicas no estilo Gran Cursos.

Lembre-se da regra editorial: nenhum markdown (**), rótulos de ícone sempre em bloco próprio ({{"tipo":"icone","nome":"..."}}),
fórmulas em {{"tipo":"formula"}}, dados tabulares em {{"tipo":"tabela"}}, diagramas mencionados pelo professor em {{"tipo":"imagem"}}.

Gere o JSON estruturado seguindo rigorosamente todas as regras editoriais. Retorne estritamente JSON válido.
"""

    print("[ProcessadorGemini] Chamando Gemini 3.6 Flash...")
    response = _cliente().models.generate_content(
        model="gemini-3.6-flash",
        contents=prompt_usuario,
        config=types.GenerateContentConfig(
            system_instruction=PROMPT_SISTEMA,
            response_mime_type="application/json",
            temperature=0.2,
        ),
    )
    
    texto_resp = response.text
    # valida json
    dados = json.loads(texto_resp)
    if isinstance(dados, dict) and "blocos" in dados:
        dados = dados["blocos"]
    if not isinstance(dados, list):
        raise ValueError("O modelo não retornou uma lista de blocos.")
    
    dados = sanitizar_blocos(dados)
    
    with open(saida_json_path, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)
        
    print(f"[ProcessadorGemini] Conteúdo estruturado salvo em {saida_json_path} ({len(dados)} blocos)")
    return dados

if __name__ == "__main__":
    if len(sys.argv) > 3:
        processar_conteudo(sys.argv[1], sys.argv[2], sys.argv[3])
    else:
        print("Uso: python processador_gemini.py <transcricao.json> <slides.json> <saida.json>")
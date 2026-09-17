# AGENTS.md — GranVic (Gran Cursos)

Pipelines de produção de materiais de estudo ("degravações") a partir de aulas em vídeo/slides.

- Plataforma: Windows + PowerShell. Antes de rodar Python no console, defina `$env:PYTHONIOENCODING="utf-8"` (o console cp1252 trava em caractéres como `∩`, `×`, acentos).
- Python: venv próprio em `_sistema/.venv`. Use paths completos via `workdir` (evite `cd`).
- Chaves: `.env` com `ASSEMBLYAI_API_KEY` e `GEMINI_API_KEY` (não versionar; nunca logar segredos).

## Arquitetura

Pipeline (ordem) em `_sistema/`:

1. `extrator_audio.py` — transcrição da aula (AssemblyAI) → `deg-<nome>.json`.
2. `extrator_slides.py` — extrai texto por página do slide PDF (PyMuPDF/pdfplumber) e **imagens/diagramas** (só os que não cobrem a maior parte da página — fundo é ignorado). Imagens salvas em `<pasta>/_imagens/`; resultado `slides.json` tem campo `imagens`.
3. `processador_gemini.py` — monta `pipeline_*_estruturado.json` com a LLM Gemini (`gemini-3.6-flash`, `response_mime_type="application/json"`). Inclui sanitização da saída (`sanitizar_blocos`): remove `**`, rótulos duplicados ("Comentário:", "Direto do Concurso", "Obs.:"→parágrafo corrido), normaliza nomes de ícones/acentos.
4. `gerador_docx.py` — aplica o estruturado sobre o template Word oficial → `material_final.docx`.
5. `auditoria.py` / `analisar_amostras.py` — compara o gerado contra a amostra/modelo via fingerprint e grava `auditoria_regras.json`.

## Blocos do estruturado

`titulo`, `corpo`, `enfase`, `icone` (`nome`), `questao`, `alternativa` (`letra a..e`), `comentario`, `citacao_lei`, `formula` (centralizada; `^`→sobrescrito, `_`→subscrito), `tabela` (`cabecalho`+`linhas`), `imagem` (`arquivo`+`legenda`), `tempo` (`minuto` — anexa "[nmin]" ao fim do último parágrafo antes do gabarito), `gabarito_fim` (`respostas`), `gabarito_comentario` (se ocorrer). Suporte a aulas de exatas (tabelas/fórmulas/imagens) já implementado; equações OMML nativas ainda não — fórmulas vão como texto centralizado.

## Auditoria — dois modos

- `mesma_aula`: o gerado e a amostra são a MESMA aula. Detecção: **título principal** idêntico (o cabeçalho da seção é a marca/sumário do curso, comum a todos os docs, e NÃO identifica a aula — ignorar). Compara tudo, inclusive quantidades/presença de ícones.
- `aula_diferente`: compara formato apenas. Ignora trechos de texto (`amostras`/`exemplos`), quantidades, e alinha ícones/caixas/citações que existem só de um lado (`alinhar_para_aula_diferente`). Usado quando a amostra é de outra aula (padrão do fluxo atual: material novo gerado no template Aula 1, comparado com amostra escolhida).

Comandos:

- `python analisar_amostras.py --comparar <doc_gerado> --amostra <amostra> [--modo mesma_aula|aula_diferente]`
- `python auditoria.py <doc_gerado> <pasta_modelos> [--modo ...]`
- `python executar.py --pasta-aula <caminho> [--amostra <docx>]`

## Convenções de formatação (aprendidas das amostras)

- Título: Arial 11, bold, CENTER, entrelinha 1.5.
- Corpo: Arial 11, JUSTIFY, entrelinha 1.5, `left_indent` ~35,4 pt + `first_line_indent` ~35,4 pt (algumas amostras usam 106,2 pt em trechos de lei — revisar por aula).
- Ícones (atencao, dica, jurisprudencia, exemplo, gabarito, comentario, direto_do_concurso, obs, pegadinha, resumo…): **Arial 11, bold, JUSTIFY, entrelinha 1.5, SEM borda**. A **linha limitadora** (borda inferior, `space=1`) vai no **último parágrafo de conteúdo** do bloco (corpo/questão/comentário/alternativa/fórmula/citação), nunca no rótulo. `gabarito_fim` também sem borda.
- Citação de lei: Arial 10, recuo ~72 pt.
- Imagem: largura 5,5 pol, legenda Arial itálico 10.
- Recuos configuráveis: `regras_estilo.json` → `estilos.corpo.recuo_esquerda_pt`/`recuo_primeira_linha_pt` (gerador lê na geração; template padrão é dentro do projeto, `Modelos Gran Cursos/Aula 1/...`).
- Aliases de ícones já vistos: alerta/aviso/cuidado→`atencao`; conceito/definicao→`dica`; destaque/detalhar/resumo→`resumo`.
- `executar.py` aceita JSONs de schema legado (sem `tipo`, campos `titulo/corpo/questao/alternativa/icone/enfase/gabarito_fim`) e normaliza automaticamente para o schema atual antes de gerar.

## Loop de aprendizado

Este arquivo é a memória do sistema: ao descobrir nova regra de formatação ou novo comportamento da LLM/Extrator, editar este AGENTS.md. Sempre que um novo material for aprovado pelo cliente, rodar `analisar_amostras.py` (inclui pastas `Entregas Gran Cursos/Aula*/` automaticamente) e consolidar as novas regras aqui e em `regras_editoriais.md`.
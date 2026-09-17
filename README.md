# GranVic — degravações Gran Cursos

Pipeline de produção de materiais de estudo ("degravações") a partir de aulas em vídeo/slides da Gran Cursos.

## Etapas (na ordem)

1. **`transcrever.py`** — transcrição do áudio da aula (AssemblyAI) → `pipeline_transcricao.json`.
2. **`extrator_slides.py`** — extrai texto por página do PDF dos slides e **imagens/diagramas** (fundo de página inteira é ignorado). Imagens vão para `<pasta>/_imagens/`.
3. **`processador_gemini.py`** — monta o `pipeline_estruturado.json` com a LLM Gemini, incluindo sanitização da saída (remove `**`, rótulos duplicados, normaliza ícones).
4. **`gerador_docx.py`** — aplica o estruturado sobre um template `.docx` oficial → `material_final.docx`. Suporta `titulo`, `corpo`, `enfase`, `icone`, `questao`, `alternativa`, `comentario`, `citacao_lei`, `formula` (`^`→sobrescrito, `_`→subscrito), `tabela`, `imagem`, `tempo`, `gabarito_fim`.
5. **`auditoria.py` / `analisar_amostras.py`** — compara o material gerado contra um modelo/amostra via fingerprint, em dois modos:
   - `mesma_aula`: mesma aula; compara tudo (inclusive quantidades).
   - `aula_diferente`: formato apenas; ignora texto, quantidades e ícones/caixas que existem só de um lado.

## Uso

```powershell
# Importante no Windows (evita erro de encoding no console cp1252):
$env:PYTHONIOENCODING = "utf-8"

python executar.py --pasta-aula <caminho> [--amostra <template.docx>]
python analisar_amostras.py --comparar <gerado.docx> --amostra <amostra.docx> [--modo mesma_aula|aula_diferente]
python auditoria.py <gerado.docx> <pasta_modelos> [--modo ...]
```

## Configuração

- Chaves de API: arquivo `.env` (copie de `.env.example`) com `ASSEMBLYAI_API_KEY` e `GEMINI_API_KEY`.
- Dependências: `pip install -r requirements.txt`.
- Regras de estilo/recuos lidas de `regras_estilo.json`; diretrizes editoriais em `regras_editoriais.md`.

## Exemplos

`transcricao_aula5.json`, `slides_aula5.json` e `estruturado_aula5.json` são dados de uma aula real (Geografia do Maranhão) usados como fixture de teste.
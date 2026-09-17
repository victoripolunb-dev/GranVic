import random
import time


class ErroRetentavel(Exception):
    """Falha transitória que deve disparar nova tentativa."""


PALAVRAS_RETENTAVEIS = (
    "429", "rate", "quota", "resource exhausted", "exhausted",
    "500", "502", "503", "504", "internal", "unavailable", "overloaded",
    "deadline", "timeout", "timed out", "temporarily", "connection",
    "reset", "broken pipe",
)


def eh_retentavel(erro):
    return isinstance(erro, ErroRetentavel) or any(p in str(erro).lower() for p in PALAVRAS_RETENTAVEIS)


def com_retry(acao, descricao, tentativas=5, base_segundos=2.0, max_segundos=60.0):
    """Executa `acao()` com backoff exponencial + jitter.

    Reexecuta apenas em falhas transitórias (rate limit, 5xx, timeout, JSON
    inválido marcado com ErroRetentavel). Erros permanentes (chave inválida,
    400 de validação) sobem na primeira tentativa.
    """
    for n in range(1, tentativas + 1):
        try:
            return acao()
        except Exception as erro:
            if n == tentativas or not eh_retentavel(erro):
                raise
            espera = min(base_segundos * (2 ** (n - 1)), max_segundos) + random.uniform(0, 0.75)
            print(f"[Retry] {descricao}: falha na tentativa {n}/{tentativas} ({erro}). "
                  f"Nova tentativa em {espera:.1f}s...")
            time.sleep(espera)

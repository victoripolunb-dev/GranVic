import os
import sys
import json
from dotenv import load_dotenv
import assemblyai as aai
from retry_utils import com_retry, ErroRetentavel

load_dotenv()
aai.settings.api_key = os.environ.get("ASSEMBLYAI_API_KEY")

def transcrever_audio(audio_path, output_json_path):
    if not os.environ.get("ASSEMBLYAI_API_KEY"):
        raise RuntimeError("ASSEMBLYAI_API_KEY não configurada no .env")
    print(f"[AssemblyAI] Enviando arquivo: {os.path.basename(audio_path)}...")
    config = aai.TranscriptionConfig(
        language_code="pt",
        punctuate=True,
        format_text=True,
    )

    def _enviar():
        transcriber = aai.Transcriber(config=config)
        transcript = transcriber.transcribe(audio_path)
        if transcript.status == aai.TranscriptStatus.error:
            raise ErroRetentavel(f"Erro na transcrição AssemblyAI: {transcript.error}")
        return transcript

    transcript = com_retry(_enviar, "transcrição AssemblyAI")

    resultado = {
        "text": transcript.text,
        "utterances": [{"start": u.start, "end": u.end, "text": u.text} for u in (transcript.utterances or [])]
    }
    
    with open(output_json_path, "w", encoding="utf-8") as f:
        json.dump(resultado, f, ensure_ascii=False, indent=2)
        
    print(f"[AssemblyAI] Transcrição concluída! Salva em {output_json_path}")
    return resultado

if __name__ == "__main__":
    if len(sys.argv) > 2:
        transcrever_audio(sys.argv[1], sys.argv[2])
    else:
        print("Uso: python transcrever.py <audio.mp3> <saida.json>")

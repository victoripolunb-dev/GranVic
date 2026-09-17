import os
import sys
import json
from dotenv import load_dotenv
import assemblyai as aai

load_dotenv()
aai.settings.api_key = os.environ.get("ASSEMBLYAI_API_KEY")

def transcrever_audio(audio_path, output_json_path):
    print(f"[AssemblyAI] Enviando arquivo: {os.path.basename(audio_path)}...")
    config = aai.TranscriptionConfig(
        language_code="pt",
        punctuate=True,
        format_text=True,
    )
    transcriber = aai.Transcriber(config=config)
    transcript = transcriber.transcribe(audio_path)
    
    if transcript.status == aai.TranscriptStatus.error:
        raise Exception(f"Erro na transcrição AssemblyAI: {transcript.error}")
        
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

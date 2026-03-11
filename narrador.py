import os
import asyncio
import re
from pptx import Presentation
from moviepy import ImageClip, AudioFileClip, concatenate_videoclips
import edge_tts

# --- CONFIGURAÇÕES ---
PPTX_FILE = "seu_arquivo.pptx"
OUTPUT_VIDEO = "resultado_final.mp4"
VOICE = "pt-BR-AntonioNeural"

def clean_text(text):
    """Remove anos e limpa espaços extras."""
    text = re.sub(r'\b\d{4}\b', '', text)
    return " ".join(text.split()).strip()

def get_all_text_from_shape(shape):
    """Busca texto recursivamente, mesmo dentro de grupos ou tabelas."""
    texts = []
    if shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            if paragraph.text.strip():
                texts.append(paragraph.text)
    
    # Se for um grupo de objetos, busca dentro de cada um deles
    if shape.shape_type == 6: # 6 é o ID para GroupShape
        for s in shape.shapes:
            texts.extend(get_all_text_from_shape(s))
            
    return texts

async def generate_audio(text, output_path):
    communicate = edge_tts.Communicate(text, VOICE)
    await communicate.save(output_path)

def extract_content(pptx_path):
    if not os.path.exists(pptx_path):
        print(f"❌ Erro: Arquivo {pptx_path} não encontrado!")
        return []

    prs = Presentation(pptx_path)
    slides_data = []
    
    for i, slide in enumerate(prs.slides):
        # 1. Prioridade Máxima: Notas do Orador
        text_notes = slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ""
        
        if text_notes.strip():
            final_text = text_notes
        else:
            # 2. Busca exaustiva em todas as formas do slide
            all_chunks = []
            # Ordena as formas pela posição (topo para baixo)
            sorted_shapes = sorted(slide.shapes, key=lambda s: (s.top, s.left))
            
            for shape in sorted_shapes:
                all_chunks.extend(get_all_text_from_shape(shape))
            
            final_text = " ".join(all_chunks)
        
        final_text = clean_text(final_text)

        if not final_text:
            final_text = "Próximo slide."

        img_path = f"slide_{i}.jpg" 
        slides_data.append({"text": final_text, "image": img_path})
    
    return slides_data

async def create_video():
    if not os.path.exists("temp"): os.makedirs("temp")
    data = extract_content(PPTX_FILE)
    clips = []

    print(f"🎙️ Iniciando narração. Total de {len(data)} slides.")

    for i, item in enumerate(data):
        audio_path = f"temp/audio_{i}.mp3"
        img_path = item['image']

        if not os.path.exists(img_path):
            print(f"⚠️ Imagem '{img_path}' não encontrada. Pulando...")
            continue

        print(f"   -> Processando slide {i} (Texto detectado: {len(item['text'])} caracteres)")
        await generate_audio(item['text'], audio_path)
        
        audio_clip = AudioFileClip(audio_path)
        video_clip = ImageClip(img_path).with_duration(audio_clip.duration)
        video_clip = video_clip.with_audio(audio_clip)
        
        clips.append(video_clip)

    if clips:
        print("🎬 Renderizando vídeo final...")
        final_video = concatenate_videoclips(clips, method="compose")
        final_video.write_videofile(OUTPUT_VIDEO, fps=24, codec="libx264")
        print(f"✅ Sucesso! Vídeo salvo como {OUTPUT_VIDEO}")

if __name__ == "__main__":
    asyncio.run(create_video())
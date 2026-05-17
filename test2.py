import io
import sounddevice as sd
import soundfile as sf
from gtts import gTTS

while True:
    text = input("\nNhập câu tiếng Việt (q để thoát): ").strip()
    if text.lower() == 'q':
        break

    print("Đang phát...")
    tts = gTTS(text=text, lang='vi')
    
    # Lưu vào buffer trong RAM, không tạo file
    buf = io.BytesIO()
    tts.write_to_fp(buf)
    buf.seek(0)
    
    data, samplerate = sf.read(buf)
    sd.play(data, samplerate)
    sd.wait()
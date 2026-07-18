import whisper
import sys
import os
import imageio_ffmpeg
import shutil

# Lấy đường dẫn ffmpeg thật
ffmpeg_exe = imageio_ffmpeg.get_ffmpeg_exe()
ffmpeg_dir = os.path.dirname(ffmpeg_exe)

# Tạo bản copy tên ffmpeg.exe nếu chưa có
ffmpeg_standard = os.path.join(ffmpeg_dir, "ffmpeg.exe")
if not os.path.exists(ffmpeg_standard):
    shutil.copy(ffmpeg_exe, ffmpeg_standard)

os.environ["PATH"] += os.pathsep + ffmpeg_dir

def transcribe_audio(audio_path: str, output_path: str = None, model_size: str = "medium"):
    """
    Trích xuất giọng nói từ file ghi âm ra file text.
    
    Args:
        audio_path: Đường dẫn đến file ghi âm (.mp3, .wav, .m4a, ...)
        output_path: Đường dẫn file text output (mặc định cùng thư mục với file ghi âm)
        model_size: Kích thước model whisper (tiny, base, small, medium, large)
    """
    if not os.path.exists(audio_path):
        print(f"❌ Không tìm thấy file: {audio_path}")
        return

    if output_path is None:
        base_name = os.path.splitext(audio_path)[0]
        output_path = base_name + ".txt"

    print(f"📂 File ghi âm : {audio_path}")
    print(f"📄 File output : {output_path}")
    print(f"🤖 Model       : {model_size}")
    print(f"⏳ Đang load model...")

    model = whisper.load_model(model_size)

    print(f"🎙️  Đang trích xuất văn bản...")
    result = model.transcribe(audio_path, language="vi")

    text = result["text"].strip()

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(text)

    print(f"✅ Hoàn thành! Đã lưu vào: {output_path}")
    print(f"\n--- NỘI DUNG ---\n{text}\n----------------")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Cách dùng: python audio_to_text.py <đường_dẫn_file_ghi_âm> [output.txt] [model_size]")
        print("Ví dụ     : python audio_to_text.py recording.mp3")
        print("Ví dụ     : python audio_to_text.py recording.mp3 output.txt large")
        print("\nModel size: tiny | base | small | medium (mặc định) | large")
        sys.exit(1)

    audio_path  = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    model_size  = sys.argv[3] if len(sys.argv) > 3 else "medium"

    transcribe_audio(audio_path, output_path, model_size)
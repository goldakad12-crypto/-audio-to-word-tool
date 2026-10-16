import streamlit as st
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import tempfile
import re
from groq import Groq

st.set_page_config(page_title="منصة التفريغ والتحليل الذكي", page_icon="🎙️", layout="wide")

st.title("🎙️ منصة التفريغ الصوتي والتحليل الذكي")
st.caption("أداة احترافية سريعة لتفريغ التسجيلات، تنظيف النصوص، والتحويل الفوري إلى Word و SRT.")

try:
    api_key = st.secrets["GROQ_API_KEY"]
except Exception:
    api_key = None

def clean_arabic_text(text):
    if not text:
        return ""
    text = re.sub(r'\s+', ' ', text).strip()
    text = re.sub(r'\s+([،.؟!؛:%])', r'\1', text)
    text = re.sub(r'([،.؟!؛:%])([^\s])', r'\1 \2', text)
    return text

def format_timestamp(seconds):
    hrs = int(seconds // 3600)
    mins = int((seconds % 3600) // 60)
    secs = int(seconds % 60)
    millis = int((seconds - int(seconds)) * 1000)
    return f"{hrs:02d}:{mins:02d}:{secs:02d},{millis:03d}"

audio_file = st.file_uploader("قم برفع الملف الصوتي أو التسجيل", type=["mp3", "wav", "m4a", "ogg", "mp4"])

if audio_file is not None:
    st.audio(audio_file)
    
    if st.button("🚀 بدء التفريغ والتحليل الفوري"):
        if not api_key:
            st.error("⚠️ تنبيه: لم يتم إعداد مفتاح الـ API في إعدادات المنصة (Secrets).")
        else:
            with st.spinner("جاري معالجة الملف الصوتي عبر السحابة الذكية..."):
                with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(audio_file.name)[1]) as tmp_file:
                    tmp_file.write(audio_file.read())
                    tmp_path = tmp_file.name

                try:
                    client = Groq(api_key=api_key)
                    with open(tmp_path, "rb") as file:
                        # استخدام النسخة السريعة والمستقرة Turbo
                        transcription = client.audio.transcriptions.create(
                            file=(audio_file.name, file.read()),
                            model="whisper-large-v3-turbo",
                            response_format="verbose_json",
                            language="ar"
                        )
                    
                    segments = transcription.segments if hasattr(transcription, 'segments') else []
                    text_full = transcription.text if hasattr(transcription, 'text') else ""

                    doc = docx.Document()
                    for section in doc.sections:
                        section.top_margin = docx.shared.Inches(1)
                        section.bottom_margin = docx.shared.Inches(1)
                        section.left_margin = docx.shared.Inches(1)
                        section.right_margin = docx.shared.Inches(1)

                    p_title = doc.add_heading("تقرير التفريغ الصوتي المنسق", level=1)
                    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    if text_full:
                        doc.add_heading("📌 ملخص وأبرز نقاط المحتوى", level=2).alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        summary_snippet = text_full[:400] + "..." if len(text_full) > 400 else text_full
                        p_sum = doc.add_paragraph(f"الملخص التمهيدي: {summary_snippet}")
                        p_sum.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        doc.add_paragraph("--------------------------------------------------")

                    doc.add_heading("📝 النص التفصيلي المفرغ", level=2).alignment = WD_ALIGN_PARAGRAPH.RIGHT

                    current_p = ""
                    word_count = 0
                    srt_content = ""
                    max_words_per_p = 60

                    if segments:
                        for idx, seg in enumerate(segments, start=1):
                            cleaned = clean_arabic_text(seg.get('text', ''))
                            words = cleaned.split()
                            start_s = seg.get('start', 0)
                            end_s = seg.get('end', 0)

                            srt_content += f"{idx}\n{format_timestamp(start_s)} --> {format_timestamp(end_s)}\n{cleaned}\n\n"

                            time_tag = f"[{int(start_s//60):02d}:{int(start_s%60):02d}]"
                            if word_count + len(words) <= max_words_per_p:
                                current_p += f" {cleaned}" if current_p else f"{time_tag} {cleaned}"
                                word_count += len(words)
                            else:
                                if current_p.strip():
                                    p = doc.add_paragraph(current_p.strip())
                                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                                current_p = f"{time_tag} {cleaned}"
                                word_count = len(words)
                    else:
                        cleaned = clean_arabic_text(text_full)
                        current_p = cleaned

                    if current_p.strip():
                        p = doc.add_paragraph(current_p.strip())
                        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

                    word_path = tempfile.NamedTemporaryFile(delete=False, suffix=".docx").name
                    doc.save(word_path)

                    st.success("✅ تم التفريغ والتحليل بنجاح تام!")

                    col1, col2 = st.columns(2)
                    with col1:
                        with open(word_path, "rb") as f:
                            st.download_button(
                                label="📄 تحميل ملف Word المنسق",
                                data=f,
                                file_name="Formatted_Transcription.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                    with col2:
                        if srt_content:
                            st.download_button(
                                label="🎬 تحميل ملف الترجمة (.SRT)",
                                data=srt_content,
                                file_name="subtitles.srt",
                                mime="text/plain"
                            )

                except Exception as e:
                    st.error(f"حدث خطأ أثناء المعالجة: {e}")
                finally:
                    if os.path.exists(tmp_path):
                        os.remove(tmp_path)

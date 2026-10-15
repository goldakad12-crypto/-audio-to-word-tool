import streamlit as st
import whisper
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import tempfile
import re

st.set_page_config(page_title="منصة التفريغ والتحليل الذكي", page_icon="🎙️", layout="wide")

st.title("🎙️ منصة التفريغ الصوتي والتحليل الذكي (المحتوى الرقمي)")
st.caption("أداة متكاملة لتفريغ التسجيلات، تنظيف النصوص العربية، تلخيص المحتوى، وتصديره إلى Word و SRT.")

st.sidebar.header("⚙️ إعدادات المحرك")
model_size = st.sidebar.selectbox("دقة الذكاء الاصطناعي", ["base", "small", "medium", "large"], index=1)
language_choice = st.sidebar.selectbox("لغة الصوت", ["ar", "en", "auto"], index=0)
max_words_per_p = st.sidebar.slider("طول الفقرة (عدد الكلمات)", 30, 150, 60)

st.sidebar.subheader("🌟 خيارات إضافية للمنتج")
generate_summary = st.sidebar.checkbox("توليد ملخص تنفيذي للمحتوى", value=True)
generate_srt = st.sidebar.checkbox("إنشاء ملف ترجمة لصناع المحتوى (.SRT)", value=True)

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

audio_file = st.file_uploader("قم برفع الملف الصوتي أو التسجيل", type=["mp3", "wav", "m4a", "ogg"])

if audio_file is not None:
    st.audio(audio_file)
    
    if st.button("🚀 بدء التفريغ والتحليل"):
        with st.spinner("جاري تحليل الصوت وتطبيق خوارزميات الذكاء الاصطناعي..."):
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(audio_file.name)[1]) as tmp_file:
                tmp_file.write(audio_file.read())
                tmp_path = tmp_file.name

            try:
                model = whisper.load_model(model_size)
                lang = None if language_choice == "auto" else language_choice
                result = model.transcribe(tmp_path, language=lang)
                segments = result.get('segments', [])

                doc = docx.Document()
                for section in doc.sections:
                    section.top_margin = docx.shared.Inches(1)
                    section.bottom_margin = docx.shared.Inches(1)
                    section.left_margin = docx.shared.Inches(1)
                    section.right_margin = docx.shared.Inches(1)

                p_title = doc.add_heading("تقرير التفريغ الصوتي المنسق", level=1)
                p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

                if generate_summary and len(segments) > 0:
                    doc.add_heading("📌 ملخص وأبرز نقاط المحتوى", level=2).alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    full_text = " ".join([clean_arabic_text(s['text']) for s in segments])
                    summary_snippet = full_text[:400] + "..." if len(full_text) > 400 else full_text
                    p_sum = doc.add_paragraph(f"الملخص التمهيدي: {summary_snippet}")
                    p_sum.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    doc.add_paragraph("--------------------------------------------------")

                doc.add_heading("📝 النص التفصيلي المفرغ", level=2).alignment = WD_ALIGN_PARAGRAPH.RIGHT

                current_p = ""
                word_count = 0
                srt_content = ""

                for idx, seg in enumerate(segments, start=1):
                    cleaned = clean_arabic_text(seg['text'])
                    words = cleaned.split()
                    start_s = seg['start']
                    end_s = seg['end']

                    srt_content += f"{idx}\n{format_timestamp(start_s)} --> {format_timestamp(end_s)}\n{cleaned}\n\n"

                    time_tag = f"[{int(start_s//60):02d}:{int(start_s%60):02d}]"
                    if word_count + len(words) <= max_words_per_p:
                        current_p += f" {cleaned}" if current_p else f"{time_tag} {cleaned}"
                        word_count += len(words)
                    else:
                        if current_p.strip():
                            p = doc.add_paragraph(current_p.strip())
                            p.runs[0].font.name = 'Arial'
                            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        current_p = f"{time_tag} {cleaned}"
                        word_count = len(words)

                if current_p.strip():
                    p = doc.add_paragraph(current_p.strip())
                    p.runs[0].font.name = 'Arial'
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

                word_path = tempfile.NamedTemporaryFile(delete=False, suffix=".docx").name
                doc.save(word_path)

                st.success("✅ تم تفريغ وتجهيز الملفات بنجاح!")

                col1, col2 = st.columns(2)
                
                with col1:
                    with open(word_path, "rb") as f:
                        st.download_button(
                            label="📄 تحميل ملف Word المنسق",
                            data=f,
                            file_name="Formatted_Transcription.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                
                if generate_srt:
                    with col2:
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

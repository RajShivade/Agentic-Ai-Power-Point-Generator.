import streamlit as st
import requests
import subprocess
import os
import sys
import traceback

# ---------- PAGE CONFIG ----------
st.set_page_config(
    page_title="Agentic PPT Generator",
    page_icon="🤖",
    layout="centered"
)

# ---------- CUSTOM CSS ----------
st.markdown("""
<style>
.main {
    background-color: #0b0f19;
}
h1, h2, h3, p, label {
    color: #e5e7eb !important;
}
textarea {
    background-color: #121826 !important;
    color: #e5e7eb !important;
    border-radius: 12px;
    border: 1px solid #2dd4bf;
}
.stButton>button {
    background: linear-gradient(90deg, #a855f7, #2dd4bf);
    color: #020617;
    font-weight: bold;
    border-radius: 14px;
    padding: 0.7em 1.2em;
    width: 100%;
}
.card {
    background-color: #121826;
    padding: 28px;
    border-radius: 18px;
    box-shadow: 0px 15px 35px rgba(168,85,247,0.25);
}
</style>
""", unsafe_allow_html=True)

# ---------- HEADER ----------
st.markdown(
    "<h1 style='text-align:center;'>🤖 Agentic AI PowerPoint Generator</h1>",
    unsafe_allow_html=True
)
# st.markdown(
#     "<p style='text-align:center; color:#9ca3af;'>Create stunning presentations using autonomous AI workflows</p>",
#     unsafe_allow_html=True
# )

st.markdown("<br>", unsafe_allow_html=True)

# ---------- CARD ----------
with st.container():
    st.markdown("<div class='card'>", unsafe_allow_html=True)

    st.subheader("📝 Describe your Presentation")
    prompt = st.text_area(
        "Example: Create a premium AI-themed PPT on Data Science with GenAI, tools and workflow",
        height=180
    )

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("🚀 Generate PowerPoint"):
        if not prompt.strip():
            st.warning("⚠️ Please enter a prompt.")
        else:
            with st.spinner("🤖 Agentic AI is designing your slides..."):
                response = requests.post(
                    url="https://rajshivade1.app.n8n.cloud/webhook-test/4859d489-e1e4-4b1f-a5a2-950596055f06",
                    json={"prompt": prompt},
                    timeout=120
                )

            if response.status_code == 200:
                st.success("✅ Presentation logic generated!")

                try:
                    raw_code = response.json().get("output", "")

                    # ---- CLEAN CODE BLOCK SAFELY ----
                    cleaned_code = (
                        raw_code
                        .replace("```python", "")
                        .replace("```", "")
                        .strip()
                    )

                    with open("app1.py", "w", encoding="utf-8") as file:
                        file.write(cleaned_code)

                    # ---- RUN GENERATED SCRIPT SAFELY ----
                    subprocess.run(
                        [sys.executable, "app1.py"],
                        check=True,
                        capture_output=True,
                        text=True
                    )

                    # ---- FIND GENERATED PPT ----
                    ppt_files = [f for f in os.listdir() if f.endswith(".pptx")]

                    if ppt_files:
                        ppt_file = ppt_files[0]

                        with open(ppt_file, "rb") as f:
                            st.download_button(
                                label="⬇️ Download PowerPoint",
                                data=f,
                                file_name="Agentic_AI_Presentation.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                            )
                    else:
                        st.error("❌ PPT file not found after generation.")

                except subprocess.CalledProcessError as e:
                    st.error("❌ Error while generating PowerPoint.")
                    st.code(e.stderr)

                except Exception as e:
                    st.error("❌ Unexpected error occurred.")
                    st.code(traceback.format_exc())

            else:
                st.error("❌ Generation failed. Please try again.")

    st.markdown("</div>", unsafe_allow_html=True)

# ---------- FOOTER ----------
st.markdown(
    "<p style='text-align:center; margin-top:40px; color:#9ca3af;'>Built by Raj Shivade • Agentic AI • Streamlit • n8n</p>",
    unsafe_allow_html=True
)

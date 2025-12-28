import streamlit as st
import requests
import subprocess
import os
import sys

st.set_page_config(page_title="AI PPT Generator", layout="centered")

st.title("AI PowerPoint Generator")
st.write("Create PowerPoint presentations using AI + n8n workflow")

st.markdown("---")

# -------- INPUT SECTION --------
st.subheader("Presentation Settings")

prompt = st.text_area(
    "Enter your presentation topic or content",
    height=150
)

slides = st.slider("Number of Slides", min_value=3, max_value=20, value=8)

font_style = st.selectbox(
    "Select Font Style",
    ["Calibri", "Arial", "Times New Roman", "Verdana", "Cambria"]
)

bg_color = st.color_picker("Choose Background Color", "#ffffff")

st.markdown("---")

generate = st.button("Generate Presentation")
if generate and prompt:
    payload = {
        "prompt": prompt,
        "slides": slides,
        "font": font_style,
        "bg_color": bg_color
    }

    with st.spinner("Generating your presentation... Please wait "):
        try:
            response = requests.post(
                url="https://romeojuliyet.app.n8n.cloud/webhook-test/4537207d-2bde-4ddc-9c8b-22ace50a20a6",
                json=payload,
                timeout=120
            )

            if response.status_code == 200:

                with open("app1.py", "w") as file:
                    code = response.json()["output"]
                    code = code.replace("```python", "").replace("```", "")
                    file.write(code)

                result = subprocess.run(
                    [sys.executable, "app1.py"],
                    capture_output=True,
                    text=True
                )

                if result.returncode == 0:
                    st.success(" PPT generated successfully!")
                else:
                    st.error(" Error while generating PPT")
                    st.code(result.stderr)

            else:
                st.error(" Failed to connect to AI workflow")

        except Exception as e:
            st.error(f" Error: {e}")

# -------- DOWNLOAD SECTION --------
st.markdown("---")
st.subheader("Download Your PowerPoint")

ppt_files = [f for f in os.listdir(os.getcwd()) if f.endswith(".pptx")]

if ppt_files:
    ppt_name = ppt_files[0] 
    ppt_path = os.path.abspath(ppt_name)

    st.success(f"Your Presentation is ready : {ppt_name}")

    with open(ppt_path, "rb") as f:
        st.download_button(
            label="Download PPT",
            data=f,
            file_name=ppt_name,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
else:
    st.info("No presentation generated yet.")


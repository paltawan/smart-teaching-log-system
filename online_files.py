"""Transfer small uploaded files through the active Streamlit WebSocket.

Vercel may route the separate HTTP upload request to another container, where
Streamlit cannot find the user's session. Component state stays on the active
WebSocket connection instead.
"""

import base64

import streamlit as st


MAX_UPLOAD_MB = 12

_file_input = st.components.v2.component(
    "teaching_file_input",
    html='<input type="file" aria-label="เลือกไฟล์" />',
    css="""
        input { font: inherit; max-width: 100%; padding: 12px; border: 1px solid #cbd5e1;
                border-radius: 10px; background: #f8fafc; }
    """,
    js="""
        export default function({data, parentElement, setStateValue}) {
            const input = parentElement.querySelector('input[type="file"]');
            input.accept = data.accept;
            input.onchange = () => {
                const file = input.files?.[0];
                if (!file) { setStateValue('file', null); return; }
                if (file.size > data.maxBytes) {
                    setStateValue('error', `ไฟล์ต้องไม่เกิน ${data.maxMb} MB`);
                    return;
                }
                const reader = new FileReader();
                reader.onload = () => {
                    setStateValue('error', null);
                    setStateValue('file', {
                        name: file.name,
                        type: file.type || 'application/octet-stream',
                        content: String(reader.result).split(',')[1],
                    });
                };
                reader.readAsDataURL(file);
            };
        }
    """,
)


class UploadedFile:
    def __init__(self, value):
        self.name = value["name"]
        self.type = value["type"]
        self._content = base64.b64decode(value["content"], validate=True)

    def getvalue(self):
        return self._content

    def read(self):
        return self._content


def file_uploader(label, extensions, key, download_url=None):
    st.write(label)
    if download_url:
        st.link_button(
            "ดาวน์โหลด template.docx",
            download_url,
            icon=":material/download:",
        )
    result = _file_input(
        key=key,
        data={
            "accept": ",".join(f".{ext}" for ext in extensions),
            "maxBytes": MAX_UPLOAD_MB * 1024 * 1024,
            "maxMb": MAX_UPLOAD_MB,
        },
        default={"file": None, "error": None},
        on_file_change=lambda: None,
        on_error_change=lambda: None,
    )
    if result.error:
        st.error(result.error)
    if not result.file:
        return None
    if not result.file["name"].lower().endswith(tuple(f".{ext}" for ext in extensions)):
        st.error("ชนิดไฟล์ไม่ถูกต้อง")
        return None
    uploaded = UploadedFile(result.file)
    if len(uploaded.getvalue()) > MAX_UPLOAD_MB * 1024 * 1024:
        st.error(f"ไฟล์ต้องไม่เกิน {MAX_UPLOAD_MB} MB")
        return None
    st.caption(f"เลือกไฟล์: {uploaded.name}")
    return uploaded

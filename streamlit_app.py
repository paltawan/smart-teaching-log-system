"""Entrypoint for the combined online app."""

import streamlit as st


st.session_state["combined_app"] = True
st.navigation(
    [
        st.Page("app.py", title="บันทึกหลังสอน", icon=":material/edit_document:", default=True),
        st.Page("project_app.py", title="โครงการสอน", icon=":material/menu_book:"),
    ],
    position="hidden",
).run()

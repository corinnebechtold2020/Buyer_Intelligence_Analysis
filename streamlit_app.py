"""Compatibility wrapper for Streamlit Cloud

Some deployment UIs expect a file named `streamlit_app.py`. This small
wrapper imports the existing `app.py` and calls its `main()` entrypoint so
the app runs the same way.
"""
from app import main


if __name__ == "__main__":
    main()

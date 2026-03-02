#!/bin/bash
python3 patch_pwa.py
exec streamlit run app.py --server.port 5000

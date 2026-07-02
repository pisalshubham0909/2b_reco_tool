@echo off
echo Starting GSTR-2B Reconciliation Tool...
python -m streamlit run reconciliation.py --server.port 8501
pause

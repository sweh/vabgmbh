IF NOT EXIST .\ms_venv (
    python -m venv ms_venv
)
CALL .\ms_venv\Scripts\activate
pip install pyinstaller xlwt paramiko scp
pyinstaller --onefile -w convert.py

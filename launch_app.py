#!/usr/bin/env python3
"""Simple launcher script that avoids Streamlit email prompt"""

import subprocess
import sys
import os

def launch_streamlit():
    """Launch Streamlit app without email prompt"""
    
    # Set environment variable to skip Streamlit email prompt
    os.environ['STREAMLIT_SERVER_HEADLESS'] = 'true'
    
    try:
        print("=== 部品リスト変換アプリ ===")
        print("起動中...")
        print("ブラウザで http://localhost:8501 にアクセスしてください")
        print("終了するには Ctrl+C を押してください")
        print("=" * 50)
        
        # Run streamlit
        subprocess.run([
            sys.executable, "-m", "streamlit", "run", 
            "parts_converter.py", 
            "--server.headless", "true",
            "--server.port", "8501",
            "--server.address", "localhost"
        ])
        
    except KeyboardInterrupt:
        print("\nアプリケーションを終了しました")
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        print("必要な依存パッケージがインストールされていることを確認してください")
        print("pip install -r requirements_simple.txt")

if __name__ == "__main__":
    launch_streamlit()

import argparse
import os
import subprocess
import sys
from pathlib import Path

from dotenv import load_dotenv

from ppt_tool.converter import PPTConverter
from ppt_tool.inspector import PPTInspector
from ppt_tool.modifier import PPTModifier


def _parse_args():
    parser = argparse.ArgumentParser(description="PPT Secretary (Gemini Powered)")
    parser.add_argument(
        "-d",
        "--debug",
        action="store_true",
        help="Print generated Python code before executing it.",
    )
    parser.add_argument(
        "-p",
        "--ppt",
        default="presentation.pptx",
        help="Target PPTX file path. If missing, it will be created automatically.",
    )
    return parser.parse_args()


def main():
    args = _parse_args()
    debug_mode = args.debug

    # 強制讀取專案根目錄的 .env
    # 假設 main.py 在 ppt_tool/main.py，根目錄就是上一層
    project_root = Path(__file__).parent.parent
    env_path = project_root / ".env"
    load_dotenv(dotenv_path=env_path)
    
    print("[INFO] PPT Secretary (Gemini Powered) Started")
    print("-----------------------------------------")
    if debug_mode:
        print("[INFO] Debug mode ON: Generated code will be printed before execution.")
    
    # 初始化模組
    converter = PPTConverter()
    inspector = PPTInspector(converter)
    modifier = PPTModifier()
    
    # 預設檔案名稱
    ppt_path = Path(args.ppt).expanduser()
    if not ppt_path.is_absolute():
        ppt_path = (Path.cwd() / ppt_path).resolve()
    current_ppt = str(ppt_path)
    
    print(f"[INFO] Target File: {current_ppt}")
    if not ppt_path.exists():
        print("[WARN] File does not exist. It will be created upon first instruction.")

    while True:
        try:
            user_input = input("\n[USER]: ").strip()
        except EOFError:
            break
            
        if not user_input:
            continue
            
        if user_input.lower() in ['exit', 'quit']:
            print("Bye! [EXIT]")
            break
        
        # 1. 如果檔案不存在，且指令是建立，則先建立空檔案
        if not ppt_path.exists():
            from pptx import Presentation
            ppt_path.parent.mkdir(parents=True, exist_ok=True)
            prs = Presentation()
            prs.save(current_ppt)
            print(f"[SUCCESS] Created new presentation: {current_ppt}")

        # 2. Inspector (眼睛)
        print("[INFO] Inspecting presentation...")
        text_summary, pdf_path = inspector.inspect(current_ppt)
        print("[INFO] Inspection finished.")
        
        # 3. Modifier (大腦 + 手)
        success, message = modifier.generate_and_execute(
            user_input, 
            text_summary, 
            pdf_path, 
            current_ppt,
            debug=debug_mode
        )
        
        if success:
            print(f"[SUCCESS] {message}")
            # 進行視覺驗證
            ok, feedback = modifier.validate_with_vision(user_input, current_ppt)
            print(f"[INFO] Validation feedback: {feedback}")
            if not ok:
                print("[WARN] Validation reported possible issues.")
            # 自動開啟 PPT 供使用者審閱
            print("[INFO] Opening PowerPoint for review...")
            try:
                if sys.platform == 'win32':
                    os.startfile(current_ppt)
                elif sys.platform == 'darwin':  # macOS
                    subprocess.run(['open', current_ppt])
                else:  # Linux
                    subprocess.run(['xdg-open', current_ppt])
            except Exception as e:
                print(f"[WARN] Could not auto-open PowerPoint: {e}")
                print(f"[INFO] Please manually open: {current_ppt}")
        else:
            print(f"[ERROR] {message}")

if __name__ == "__main__":
    main()

import os
import sys
import subprocess
import platform
from pathlib import Path

class PPTConverter:
    def __init__(self):
        self.engine = self._detect_engine()
        print(f"[INFO] PPT Converter Engine: {self.engine}")

    def _detect_engine(self):
        """偵測可用的轉檔引擎"""
        # 1. 優先嘗試 Windows COM (僅限 Windows 且有安裝 PowerPoint)
        if sys.platform == 'win32':
            try:
                import win32com.client
                # 嘗試初始化 PowerPoint Application 檢查是否存在
                ppt_app = win32com.client.Dispatch("PowerPoint.Application")
                ppt_app.Quit()
                return "com"
            except Exception:
                pass

        # 2. 嘗試 LibreOffice (跨平台)
        # 常見的 LibreOffice 路徑或指令
        soffice_cmds = ["soffice", "libreoffice"]
        if sys.platform == 'darwin':
            soffice_cmds.append("/Applications/LibreOffice.app/Contents/MacOS/soffice")
        elif sys.platform == 'win32':
            soffice_cmds.append(r"C:\Program Files\LibreOffice\program\soffice.exe")

        for cmd in soffice_cmds:
            try:
                subprocess.run([cmd, "--version"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
                self.soffice_cmd = cmd
                return "libreoffice"
            except (FileNotFoundError, subprocess.CalledProcessError):
                continue
        
        return "none"

    def convert_to_pdf(self, ppt_path: str, output_dir: str) -> str:
        """將 PPT 轉換為 PDF，回傳 PDF 路徑"""
        ppt_path = str(Path(ppt_path).resolve())
        output_dir = str(Path(output_dir).resolve())
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        filename = Path(ppt_path).stem
        pdf_path = os.path.join(output_dir, f"{filename}.pdf")

        if self.engine == "com":
            return self._convert_with_com(ppt_path, pdf_path)
        elif self.engine == "libreoffice":
            return self._convert_with_libreoffice(ppt_path, output_dir)
        else:
            print("[WARN] No conversion engine found. Skipping visual inspection.")
            return None

    def _convert_with_com(self, ppt_path, pdf_path):
        import win32com.client
        ppt_app = None
        pres = None
        try:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            pres = ppt_app.Presentations.Open(ppt_path, WithWindow=False)
            
            # 檢查是否有投影片，空簡報無法轉 PDF
            if pres.Slides.Count == 0:
                print("[WARN] Presentation has no slides, skipping PDF conversion.")
                return None
            
            pres.SaveAs(pdf_path, 32) # 32 = ppSaveAsPDF
            return pdf_path
        except Exception as e:
            print(f"[ERROR] COM Conversion Error: {e}")
            return None
        finally:
            if pres:
                pres.Close()
            # 不關閉 Application，以免影響使用者正在開啟的其他 PPT
            # if ppt_app: ppt_app.Quit() 

    def _convert_with_libreoffice(self, ppt_path, output_dir):
        try:
            # soffice --headless --convert-to pdf <file> --outdir <dir>
            cmd = [
                self.soffice_cmd,
                "--headless",
                "--convert-to", "pdf",
                ppt_path,
                "--outdir", output_dir
            ]
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            
            filename = Path(ppt_path).stem
            expected_pdf = os.path.join(output_dir, f"{filename}.pdf")
            if os.path.exists(expected_pdf):
                return expected_pdf
            return None
        except Exception as e:
            print(f"[ERROR] LibreOffice Conversion Error: {e}")
            return None

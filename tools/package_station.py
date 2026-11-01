r"""Đóng gói bản phát hành Trạm SAP: dist/TramSAP-<phiên bản>.zip

    python3 tools/package_station.py            # phiên bản lấy từ "git describe"
    python3 tools/package_station.py 23.1.3     # đặt phiên bản

Gói gồm sap-station/ (trừ cấu hình thật), sap-simulator/, tools/vba/ và file VERSION.
Trên máy SAP: giải nén, chạy sap-station\install.ps1 (xem sap-station/docs/HUONG_DAN_CAI_TRAM_SAP.txt).
"""
import os, re, subprocess, sys, zipfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
INCLUDE = ["sap-station", "sap-simulator", os.path.join("tools", "vba")]
EXCLUDE = {os.path.join("sap-station", "TramSAP.config.json")}


def version():
    if len(sys.argv) > 1:
        return sys.argv[1]
    try:
        v = subprocess.run(["git", "describe", "--tags", "--always", "--dirty"], cwd=ROOT, capture_output=True, text=True, check=True).stdout.strip()
    except Exception:
        v = "dev"
    return v


def macro_version():
    s = open(os.path.join(ROOT, "sap-station", "vba", "SAP_Automation.bas"), encoding="latin-1").read()
    m = re.search(r'MACRO_VERSION As String = "([^"]+)"', s)
    return m.group(1) if m else "?"


def main():
    v = version()
    os.makedirs(os.path.join(ROOT, "dist"), exist_ok=True)
    out = os.path.join(ROOT, "dist", f"TramSAP-{v}.zip")
    n = 0
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("VERSION", f"{v}\r\nmacro {macro_version()}\r\n")
        for top in INCLUDE:
            for d, _, files in os.walk(os.path.join(ROOT, top)):
                for f in sorted(files):
                    full = os.path.join(d, f)
                    rel = os.path.relpath(full, ROOT)
                    if rel in EXCLUDE or "__pycache__" in rel:
                        continue
                    z.write(full, rel.replace(os.sep, "/")); n += 1
    print(f"Đã tạo {os.path.relpath(out, ROOT)} ({n} file, macro {macro_version()})")


if __name__ == "__main__":
    main()

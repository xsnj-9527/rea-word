# -*- coding: utf-8 -*-
"""Build DocPreview: exe -> installer -> portable zip.

Run with:  python build.py
Requires PyInstaller and Inno Setup 6 on this machine.
"""

import os
import re
import shutil
import subprocess
import sys
import zipfile

ROOT = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable
ISCC_CANDIDATES = [
    os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Programs\Inno Setup 6\ISCC.exe"),
    r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    r"C:\Program Files\Inno Setup 6\ISCC.exe",
]


def read_version():
    with open(os.path.join(ROOT, "src", "app.py"), encoding="utf-8") as f:
        m = re.search(r'^VERSION\s*=\s*"([^"]+)"', f.read(), re.M)
    if not m:
        raise SystemExit("无法从 src/app.py 读取 VERSION")
    return m.group(1)


def run(cmd, **kw):
    print("  $ " + " ".join(cmd) if isinstance(cmd, list) else "  $ " + cmd)
    return subprocess.run(cmd, cwd=ROOT, **kw)


def main():
    version = read_version()
    print("版本:", version)

    print("\n=== 1/4 清理旧产物 ===")
    for d in ("build", "dist", "installer_out"):
        p = os.path.join(ROOT, d)
        if os.path.isdir(p):
            shutil.rmtree(p)
            print("  已删除", d)

    print("\n=== 2/4 打包 exe ===")
    r = run([PY, "-m", "PyInstaller", "DocPreview.spec", "--noconfirm", "--clean"])
    if r.returncode != 0:
        raise SystemExit("PyInstaller 失败")
    exe = os.path.join(ROOT, "dist", "DocPreview.exe")
    if not os.path.isfile(exe):
        raise SystemExit("未生成 dist/DocPreview.exe")
    print("  exe: %.1f MB" % (os.path.getsize(exe) / 1048576))

    print("\n=== 3/4 生成安装包 ===")
    iscc = next((p for p in ISCC_CANDIDATES if os.path.isfile(p)), None)
    if not iscc:
        print("  [跳过] 未找到 Inno Setup")
    else:
        r = run([iscc, "/DAppVersion=%s" % version, "installer.iss"])
        if r.returncode != 0:
            raise SystemExit("Inno Setup 失败")

    print("\n=== 4/4 生成绿色版 ===")
    zip_path = os.path.join(ROOT, "dist", "DocPreview-%s-portable.zip" % version)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(exe, "DocPreview.exe")
        for extra in ("README.md", "LICENSE"):
            p = os.path.join(ROOT, extra)
            if os.path.isfile(p):
                z.write(p, extra)
    print("  zip: %.1f MB" % (os.path.getsize(zip_path) / 1048576))

    print("\n=== 完成 ===")
    for folder in ("dist", "installer_out"):
        d = os.path.join(ROOT, folder)
        if not os.path.isdir(d):
            continue
        for f in sorted(os.listdir(d)):
            p = os.path.join(d, f)
            if os.path.isfile(p):
                print("  %-45s %6.1f MB" % (os.path.join(folder, f),
                                            os.path.getsize(p) / 1048576))


if __name__ == "__main__":
    main()

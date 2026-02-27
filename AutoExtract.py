#!/usr/bin/env python3
"""
AutoExtract — Archive Extraction Tool
"""

import os
import shutil
from pathlib import Path
import sys
from collections import defaultdict
import subprocess
import platform
from datetime import datetime

# Terminal Colors
R  = "\033[0m"
B  = "\033[1m"
D  = "\033[2m"
IT = "\033[3m"

# Warm amber/orange palette — feels like unpacking treasure
AMBER  = "\033[38;5;214m"
GOLD   = "\033[38;5;220m"
ORANGE = "\033[38;5;208m"
CREAM  = "\033[38;5;230m"
BROWN  = "\033[38;5;130m"
TEAL   = "\033[38;5;73m"
SAGE   = "\033[38;5;108m"
ROSE   = "\033[38;5;210m"
MUTED  = "\033[38;5;245m"
WHITE  = "\033[97m"
RED    = "\033[91m"
GREEN  = "\033[92m"

BG_DARK = "\033[48;5;234m"

def clear():
    os.system("cls" if os.name == "nt" else "clear")

# Palette warna tetap sama biar matching sama setup kamu
AMBER  = "\033[38;5;214m"
GOLD   = "\033[38;5;220m"
CREAM  = "\033[38;5;230m"
B      = "\033[1m"
DIM    = "\033[2m"
R      = "\033[0m"

def banner():
    print(f"""
{AMBER}{B}  ╭─ {GOLD}AUTOEXTRACT {CREAM}v1.0{AMBER} ────────────────────────────╮
  │  {GOLD}█▀█ █░█ ▀█▀ █▀█ █▀▀ ▀▄▀ ▀█▀ █▀█ █▀█ █▀▀ ▀█▀{AMBER}  │
  │  {GOLD}█▀█ █▄█ ░█░ █▄█ ██▄ █░█ ░█░ █▀▄ █▀█ █▄▄ ░█░{AMBER}  │
  │  {DIM}───────────────────────────────────────────{R}{AMBER}  │
  │  {CREAM}Smarter Archive Unpacker & File Manager{AMBER}      │
  ╰───────────────────────────────────────────────╯{R}
""")

def rule(char="─", n=65, color=MUTED):
    print(f"{color}{char * n}{R}")

def header(text, icon="◈"):
    print(f"\n{GOLD}{B}{icon} {text}{R}")
    rule("╌", color=AMBER)

def fmt_size(b):
    for u in ("B","KB","MB","GB","TB"):
        if b < 1024: return f"{b:.1f} {u}"
        b /= 1024
    return f"{b:.1f} PB"

def badge(text, color=TEAL):
    return f"{color}{B} {text} {R}"

def pause(msg="↵  Press Enter to continue"):
    input(f"\n{MUTED}{IT}{msg}{R}  ")

def spin_progress(i, total, label=""):
    frames = ["⠋","⠙","⠹","⠸","⠼","⠴","⠦","⠧","⠇","⠏"]
    frame = frames[i % len(frames)]
    pct = int(i / total * 30)
    bar = f"{AMBER}{'━' * pct}{MUTED}{'─' * (30 - pct)}{R}"
    print(f"  {GOLD}{frame}{R}  [{bar}]  {MUTED}{i}/{total}  {D}{label[:28]}{R}", end="\r")


# ArchiveExtractor
class ArchiveExtractor:
    SUPPORTED = {'.zip','.rar','.7z','.tar','.gz','.bz2','.xz',
                 '.tar.gz','.tar.bz2','.tar.xz'}

    def __init__(self):
        self.results    = defaultdict(list)
        self.total      = 0
        self.success    = 0
        self.failed     = 0
        self.pw_count   = 0
        self.global_pw  = None
        self.extracted  = []

    def is_archive(self, fp):
        return Path(fp).suffix.lower() in self.SUPPORTED

    def find_archives(self, directory, recursive):
        found = []
        if recursive:
            for root, _, files in os.walk(directory):
                for f in files:
                    fp = os.path.join(root, f)
                    if self.is_archive(fp): found.append(fp)
        else:
            for item in os.listdir(directory):
                fp = os.path.join(directory, item)
                if os.path.isfile(fp) and self.is_archive(fp): found.append(fp)
        return sorted(found)

    def get_out_path(self, archive_path):
        base = os.path.join(os.path.dirname(archive_path), Path(archive_path).stem)
        path, n = base, 1
        while os.path.exists(path): path = f"{base}_{n}"; n += 1
        return path

    def collect(self, out_path):
        files, total = [], 0
        for root, _, filenames in os.walk(out_path):
            for f in filenames:
                fp = os.path.join(root, f)
                try:
                    sz = os.path.getsize(fp)
                    files.append({'path': fp, 'size': sz,
                                  'rel': os.path.relpath(fp, out_path)})
                    total += sz
                except OSError: pass
        return files, total

    def find_7zip(self):
        paths = (["C:\\Program Files\\7-Zip\\7z.exe","7z.exe","7z"]
                 if platform.system()=="Windows"
                 else ["/usr/bin/7z","/usr/bin/7za","/usr/local/bin/7z","7z","7za"])
        for p in paths:
            if os.path.exists(p): return p
            try:
                r = subprocess.run([p,"--help"],capture_output=True,timeout=5,check=False)
                if r.returncode in (0,1): return p
            except: pass
        return None

    def extract_7zip(self, archive, out, password=None):
        exe = self.find_7zip()
        if not exe: return False, "7-Zip not found"
        cmd = [exe, "x", archive, f"-o{out}", "-y"]
        if password: cmd.append(f"-p{password}")
        try:
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
            if r.returncode == 0: return True, "ok"
            err = (r.stderr + r.stdout).lower()
            if "wrong password" in err or "encrypted" in err:
                self.pw_count += 1
                return False, "Wrong password or encrypted"
            return False, f"7z error code {r.returncode}"
        except subprocess.TimeoutExpired: return False, "Timed out"
        except Exception as e: return False, str(e)

    def extract_patool(self, archive, out, password=None):
        try:
            import patoolib
            os.makedirs(out, exist_ok=True)
            kw = {"password": password} if password else {}
            try:
                patoolib.extract_archive(archive, outdir=out, **kw)
                return True, "ok"
            except patoolib.util.PatoolError as e:
                msg = str(e)
                if "password" in msg.lower() or "encrypted" in msg.lower():
                    self.pw_count += 1
                    return False, "Password required or incorrect"
                return False, msg
        except Exception as e: return False, str(e)

    def extract_one(self, archive_path, pw_policy, current_pw=None):
        name    = os.path.basename(archive_path)
        out     = self.get_out_path(archive_path)
        size    = os.path.getsize(archive_path) if os.path.exists(archive_path) else 0

        print(f"\n  {AMBER}▸{R} {WHITE}{B}{name}{R}  {MUTED}({fmt_size(size)}){R}")
        print(f"    {MUTED}⤷ {out}{R}")

        pw, max_tries, tries = current_pw, 3, 0
        ok, msg = False, "Extraction failed"

        while tries < max_tries:
            if pw_policy == "ask_each" and not pw:
                ask = input(f"    {TEAL}⚿  Password required? [y/N]: {R}").strip().lower()
                if ask == "y":
                    pw = input(f"    {TEAL}⚿  Password: {R}").strip()
                    self.pw_count += 1

            ok, msg = self.extract_7zip(archive_path, out, pw)
            if not ok and "not found" in msg:
                ok, msg = self.extract_patool(archive_path, out, pw)

            if ok: break
            if "password" in msg.lower() and pw_policy=="ask_each" and tries < max_tries-1:
                print(f"    {ROSE}✗  {msg}{R}")
                pw = input(f"    {TEAL}⚿  Retry password ({tries+2}/{max_tries}): {R}").strip()
                tries += 1
            else: break

        if not ok:
            try:
                if os.path.exists(out) and not os.listdir(out): os.rmdir(out)
            except OSError: pass

        rec = {"path": archive_path, "ok": ok, "msg": msg,
               "size": size, "out": out if ok else None}

        if ok:
            files, total_sz = self.collect(out)
            rec.update({"files": files, "ext_size": total_sz, "count": len(files)})
            self.extracted.extend(files)
            self.success += 1
            print(f"    {GREEN}{B}✓{R}  {SAGE}Extracted {len(files)} file(s)  ·  {fmt_size(total_sz)}{R}")
            self.results["ok"].append(rec)
        else:
            self.failed += 1
            print(f"    {RED}✗{R}  {ROSE}{msg}{R}")
            self.results["fail"].append(rec)

        self.results["all"].append(rec)
        return ok

    def extract_all(self, directory, recursive, pw_policy):
        archives = self.find_archives(directory, recursive)
        self.total = len(archives)

        if not archives:
            print(f"\n  {ROSE}No supported archive files found here.{R}")
            return

        header(f"EXTRACTING  ·  {self.total} archive(s)", "◈")
        pw = self.global_pw if pw_policy == "use_global" else None

        for i, ap in enumerate(archives, 1):
            print(f"\n  {MUTED}[ {i} / {self.total} ]{R}", end="")
            self.extract_one(ap, pw_policy, pw)

    # Menus
    def menu_scan_mode(self):
        header("SCAN MODE", "◈")
        cwd = os.getcwd()
        print(f"  {MUTED}Location  {AMBER}{cwd}{R}\n")
        print(f"  {GOLD}1{R}  {CREAM}Current folder only{R}  {MUTED}(skip subfolders){R}")
        print(f"  {GOLD}2{R}  {CREAM}All folders recursively{R}  {MUTED}(scan everything){R}")
        rule("╌", color=AMBER)
        while True:
            c = input(f"\n  {TEAL}→  {R}").strip()
            if c == "1": return False
            if c == "2": return True
            print(f"  {ROSE}Enter 1 or 2.{R}")

    def menu_password_policy(self):
        header("PASSWORD POLICY", "◈")
        print(f"  {MUTED}How to handle encrypted archives?\n{R}")
        print(f"  {GOLD}1{R}  {CREAM}Ask per archive{R}")
        print(f"  {GOLD}2{R}  {CREAM}One global password{R}")
        print(f"  {GOLD}3{R}  {CREAM}Skip encrypted files{R}")
        rule("╌", color=AMBER)
        while True:
            c = input(f"\n  {TEAL}→  {R}").strip()
            if c == "1": return "ask_each"
            if c == "2":
                self.global_pw = input(f"  {TEAL}⚿  Global password: {R}").strip()
                return "use_global"
            if c == "3": return "skip_all"
            print(f"  {ROSE}Enter 1, 2, or 3.{R}")

    def menu_copy(self):
        if not self.extracted: return
        total_sz = sum(f["size"] for f in self.extracted)
        header(f"COPY FILES  ·  {len(self.extracted)} file(s)  ·  {fmt_size(total_sz)}", "◈")
        print(f"  {GOLD}1{R}  {CREAM}Copy all to a folder{R}")
        print(f"  {GOLD}2{R}  {CREAM}Pick specific files{R}")
        print(f"  {GOLD}3{R}  {CREAM}Leave in place{R}")
        rule("╌", color=AMBER)
        while True:
            c = input(f"\n  {TEAL}→  {R}").strip()
            if c == "1": self._do_copy(self.extracted); return
            if c == "2": self._selective(); return
            if c == "3":
                print(f"\n  {SAGE}✓  Files remain in extraction folders.{R}")
                return
            print(f"  {ROSE}Enter 1, 2, or 3.{R}")

    def _ask_target(self):
        t = input(f"  {TEAL}⤷  Target folder: {R}").strip().strip('"').strip("'")
        if not os.path.exists(t):
            c = input(f"  {AMBER}Doesn't exist — create it? [y/N]: {R}").strip().lower()
            if c == "y":
                try: os.makedirs(t, exist_ok=True); print(f"  {SAGE}✓  Created.{R}"); return t
                except OSError as e: print(f"  {ROSE}✗  {e}{R}"); return None
            return None
        if not os.path.isdir(t): print(f"  {ROSE}Not a directory.{R}"); return None
        return t

    def _do_copy(self, file_list):
        target = self._ask_target()
        if not target: return
        copied, skipped, size = 0, 0, 0
        n = len(file_list)
        print()
        for i, fi in enumerate(file_list, 1):
            src  = fi["path"]
            name = os.path.basename(src)
            dst  = os.path.join(target, name)
            k = 1
            while os.path.exists(dst):
                stem, ext = os.path.splitext(name)
                dst = os.path.join(target, f"{stem}_{k}{ext}"); k += 1
            spin_progress(i, n, name)
            try:
                shutil.copy2(src, dst)
                copied += 1; size += fi["size"]
            except Exception: skipped += 1
        print(" " * 72, end="\r")
        print(f"  {GREEN}✓{R}  {SAGE}Copied {copied}/{n} file(s)  ·  {fmt_size(size)}{R}")
        if skipped: print(f"  {ROSE}✗  {skipped} failed{R}")
        print(f"  {MUTED}⤷ {target}{R}")

    def _selective(self):
        header(f"SELECT FILES  ·  {len(self.extracted)} available", "◈")
        for i, fi in enumerate(self.extracted, 1):
            name = os.path.basename(fi["path"])
            print(f"  {GOLD}{i:>3}{R}  {CREAM}{name}{R}  {MUTED}{fmt_size(fi['size'])}{R}")
        rule("╌", color=AMBER)
        print(f"  {MUTED}Numbers separated by commas, or 'all'{R}\n")
        while True:
            raw = input(f"  {TEAL}→  {R}").strip().lower()
            if raw == "all": self._do_copy(self.extracted); return
            try:
                sel = [self.extracted[int(x.strip())-1] for x in raw.split(",")
                       if x.strip() and 0 < int(x.strip()) <= len(self.extracted)]
                if sel: self._do_copy(sel); return
                print(f"  {ROSE}No valid files selected.{R}")
            except (ValueError, IndexError):
                print(f"  {ROSE}Invalid input.{R}")

    def show_summary(self):
        header("SUMMARY", "◈")
        total_in = sum(r["size"] for r in self.results["all"])
        total_out = sum(r.get("ext_size",0) for r in self.results["ok"])
        total_f = sum(r.get("count",0) for r in self.results["ok"])
        rate = (self.success / self.total * 100) if self.total else 0

        col_w = 24
        rows = [
            ("Archives found",    str(self.total)),
            ("Extracted",         f"{GREEN}{B}{self.success}{R}"),
            ("Failed",            f"{(RED if self.failed else SAGE)}{self.failed}{R}"),
            ("Encrypted",         str(self.pw_count)),
            ("Input size",        fmt_size(total_in)),
            ("Output size",       fmt_size(total_out)),
            ("Files unpacked",    str(total_f)),
            ("Success rate",      f"{rate:.0f}%"),
        ]
        for label, val in rows:
            print(f"  {MUTED}{label:<{col_w}}{R}{CREAM}{val}{R}")

        if self.results["fail"]:
            print(f"\n  {ROSE}Failed:{R}")
            for r in self.results["fail"]:
                print(f"  {MUTED}  ·  {os.path.basename(r['path'])} — {r['msg']}{R}")

        if self.results["ok"]:
            print(f"\n  {SAGE}Extracted:{R}")
            for r in self.results["ok"]:
                print(f"  {MUTED}  ·  {os.path.basename(r['path'])}  ⤷  {r.get('count',0)} file(s){R}")

        rule()
        if self.success > 0:
            print(f"\n  {AMBER}{B}All done! 🎉{R}")
        else:
            print(f"\n  {MUTED}Nothing was extracted.{R}")


def check_deps():
    try: import patoolib; return True
    except ImportError:
        print(f"\n  {ROSE}Missing: patool{R}  →  {MUTED}pip install patool{R}")
        return False

def main():
    clear(); banner()
    cwd = os.getcwd()
    print(f"  {MUTED}Platform   {AMBER}{platform.system()}{R}")
    print(f"  {MUTED}Directory  {AMBER}{cwd}{R}")
    print(f"  {MUTED}Formats    {AMBER}ZIP · RAR · 7Z · TAR · GZ · BZ2 · XZ{R}")
    print(f"  {MUTED}Engine     {AMBER}7-Zip → patool{R}")

    ex = ArchiveExtractor()
    try:
        recursive = ex.menu_scan_mode()
        pw_policy = ex.menu_password_policy()
        clear(); banner()
        ex.extract_all(cwd, recursive, pw_policy)
        ex.menu_copy()
        clear(); banner()
        ex.show_summary()
    except KeyboardInterrupt:
        print(f"\n\n  {AMBER}Interrupted.{R}")
        ex.show_summary()
    except Exception as e:
        print(f"\n  {ROSE}Error: {e}{R}"); sys.exit(1)

if __name__ == "__main__":
    if check_deps(): main()

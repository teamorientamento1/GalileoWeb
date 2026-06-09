#!/usr/bin/env python3
"""
carica_attestati.py
───────────────────
Legge l'Excel degli iscritti, ricava i CF del SummerCamp,
controlla che nella cartella indicata ci siano i PDF corrispondenti
e li copia nel repo GalileoWeb / fa il push su GitHub.

USO:
    python carica_attestati.py --cartella /percorso/ai/tuoi/pdf
    python carica_attestati.py --cartella ~/Desktop/attestati --dry-run
    python carica_attestati.py --cartella . --no-push   # copia ma non fa push

PRESUPPONE che i PDF siano nominati esattamente come il CF, es:
    RSSMRA90A01H501Z.pdf
    GRDSRA07C59E463A.pdf
    ...
"""

import argparse
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    sys.exit("❌  Installa pandas: pip install pandas openpyxl")

# ── Configurazione ─────────────────────────────────────────────────────────────
EXCEL_FILE = "data/iscritti.xlsx"
EVENT_COL  = (
    "A quale evento vuoi partecipare oppure a quale evento stai partecipando?\n"
    "NB. se l'evento non è presente, si prega di inserire nome dipartimento e data\n"
)
SC11 = "SummerCamp (11 giugno)"
SC12 = "SummerCamp (12 giugno)"
CF_COL = "CodiceFiscale"

# Path del repo GalileoWeb (cartella attestati/)
REPO_ATTESTATI = Path(__file__).parent.parent / "GalileoWeb" / "attestati"
# Se il repo è altrove, cambia questo:
# REPO_ATTESTATI = Path("/percorso/assoluto/GalileoWeb/attestati")

CF_REGEX = re.compile(
    r'^[A-Z]{6}[0-9LMNPQRSTUV]{2}[A-Z][0-9LMNPQRSTUV]{2}[A-Z][0-9LMNPQRSTUV]{3}[A-Z]$'
)

# ──────────────────────────────────────────────────────────────────────────────

def get_cf_summercamp(excel: Path) -> dict[str, str]:
    """Ritorna {CF_normalizzato: nome} per tutti gli iscritti al SummerCamp."""
    df  = pd.read_excel(excel, dtype=str)
    sc  = df[df[EVENT_COL].isin([SC11, SC12])].copy()
    sc["_cf"] = sc[CF_COL].str.strip().str.upper()

    cf_nome = {}
    for _, row in sc.iterrows():
        cf   = str(row["_cf"]).strip()
        nome = f"{str(row.get('Nome2','')).strip()} {str(row.get('Cognome','')).strip()}".strip()
        if cf and cf.lower() not in ("nan",""):
            cf_nome[cf] = nome   # ultima riga vince (dedup)
    return cf_nome


def run_git(repo: Path, *args):
    result = subprocess.run(["git"] + list(args), cwd=repo, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"  ⚠️  git {' '.join(args)} → {result.stderr.strip()}")
    return result.returncode == 0


def main():
    parser = argparse.ArgumentParser(description="Carica attestati PDF su GalileoWeb")
    parser.add_argument("--cartella", required=True,
                        help="Cartella con i PDF nominati {CF}.pdf")
    parser.add_argument("--excel",    default=EXCEL_FILE,
                        help=f"File Excel iscritti (default: {EXCEL_FILE})")
    parser.add_argument("--repo",     default=str(REPO_ATTESTATI),
                        help="Cartella attestati/ nel repo GalileoWeb")
    parser.add_argument("--dry-run",  action="store_true",
                        help="Mostra cosa farebbe senza copiare nulla")
    parser.add_argument("--no-push",  action="store_true",
                        help="Copia i file ma non fa git push")
    args = parser.parse_args()

    src_dir  = Path(args.cartella).expanduser().resolve()
    dest_dir = Path(args.repo).expanduser().resolve()
    excel    = Path(args.excel).expanduser()

    # ── Verifica percorsi ────────────────────────────────────────────────────
    if not src_dir.is_dir():
        sys.exit(f"❌  Cartella non trovata: {src_dir}")
    if not excel.exists():
        sys.exit(f"❌  File Excel non trovato: {excel}")

    # ── Carica CF attesi dall'Excel ──────────────────────────────────────────
    print(f"📖 Leggo Excel: {excel}")
    cf_attesi = get_cf_summercamp(excel)
    print(f"   → {len(cf_attesi)} CF univoci per il SummerCamp")

    # ── Scansiona cartella sorgente ──────────────────────────────────────────
    pdf_trovati = {
        f.stem.upper(): f
        for f in src_dir.glob("*.pdf")
        if f.stem.upper() != f.stem.upper().startswith(".")  # escludi nascosti
    }
    print(f"\n📂 PDF trovati in {src_dir}: {len(pdf_trovati)}")

    # ── Incrocia: CF attesi vs PDF trovati ───────────────────────────────────
    cf_ok      = [cf for cf in cf_attesi if cf in pdf_trovati]
    cf_mancanti= [cf for cf in cf_attesi if cf not in pdf_trovati]
    pdf_extra  = [cf for cf in pdf_trovati if cf not in cf_attesi]

    print(f"\n✅ PDF pronti da caricare : {len(cf_ok)}")
    if cf_mancanti:
        print(f"⚠️  PDF MANCANTI ({len(cf_mancanti)}):")
        for cf in sorted(cf_mancanti):
            print(f"   ✗ {cf}.pdf  ({cf_attesi[cf]})")
    if pdf_extra:
        print(f"ℹ️  PDF extra non in lista ({len(pdf_extra)}) — ignorati:")
        for cf in sorted(pdf_extra)[:10]:
            print(f"   ? {cf}.pdf")
        if len(pdf_extra) > 10:
            print(f"   … e altri {len(pdf_extra)-10}")

    if not cf_ok:
        sys.exit("\n❌ Nessun PDF da caricare. Controlla la cartella.")

    if args.dry_run:
        print(f"\n[dry-run] Verrebbero copiati {len(cf_ok)} file in {dest_dir}")
        print("[dry-run] Nessun file modificato.")
        return

    # ── Copia PDF nella cartella del repo ────────────────────────────────────
    dest_dir.mkdir(parents=True, exist_ok=True)
    copiati = 0
    for cf in cf_ok:
        src  = pdf_trovati[cf]
        dest = dest_dir / f"{cf}.pdf"
        shutil.copy2(src, dest)
        copiati += 1

    print(f"\n📁 Copiati {copiati} PDF in: {dest_dir}")

    if args.no_push:
        print("✅ File copiati. Push non eseguito (--no-push).")
        return

    # ── Git add + commit + push ──────────────────────────────────────────────
    repo_root = dest_dir.parent  # GalileoWeb/
    if not (repo_root / ".git").exists():
        print(f"⚠️  {repo_root} non è un repo git. Copia eseguita, push saltato.")
        print("   Esegui manualmente: git -C {repo_root} add attestati/ && git push")
        return

    print("\n🚀 Git push…")
    run_git(repo_root, "add", "attestati/")

    msg = f"Carica {copiati} attestati SummerCamp 2026"
    if cf_mancanti:
        msg += f" ({len(cf_mancanti)} ancora mancanti)"
    run_git(repo_root, "commit", "-m", msg)

    ok = run_git(repo_root, "push")
    if ok:
        print(f"✅ Push completato! Gli attestati sono online.")
        print(f"   Link studenti: https://teamorientamento1.github.io/GalileoWeb/attestati.html")
    else:
        print("⚠️  Push fallito. Esegui manualmente: git -C {repo_root} push")

    # ── Riepilogo finale ─────────────────────────────────────────────────────
    print(f"\n{'─'*50}")
    print(f"  Caricati  : {copiati} / {len(cf_attesi)}")
    if cf_mancanti:
        print(f"  Mancanti  : {len(cf_mancanti)} (aggiungi i PDF e riesegui)")
    print(f"{'─'*50}")


if __name__ == "__main__":
    main()

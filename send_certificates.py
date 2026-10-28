#!/usr/bin/env python3
"""Safely match existing MLSA PDF certificates to the registration workbook and email via Gmail.

Default: dry run only; writes report.csv and does not send mail.
Test:    python send_certificates.py --test-email you@example.com
Live:    python send_certificates.py --send   (requires typing SEND at the prompt)
"""
from __future__ import annotations
import argparse, csv, getpass, re, shutil, smtplib, ssl, sys, unicodedata, zipfile
from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path
import pandas as pd
from rapidfuzz import fuzz

BASE = Path(__file__).resolve().parent
EXCEL_FILE = BASE / "students.xlsx"
ZIP_FILE = BASE / "certificates.zip"
PDF_DIR = BASE / "certificates"
REPORT_FILE = BASE / "report.csv"
LOG_FILE = BASE / "send_log.csv"
SENDER = "mlsadypcet@gmail.com"
NAME_COLUMN = "Full name"
EMAIL_COLUMN = "Email2"
MIN_SCORE = 88
MIN_GAP = 5
GMAIL_HOST, GMAIL_PORT = "smtp.gmail.com", 465
REPORT_FIELDS = ["excel_row", "student_name", "email", "selected_pdf", "match_score", "status", "details"]
LOG_FIELDS = ["timestamp", "student_name", "email", "pdf_path", "status", "details"]


def normalize(value: str) -> str:
    value = unicodedata.normalize("NFKD", str(value or ""))
    value = "".join(c for c in value if not unicodedata.combining(c)).lower()
    return " ".join(re.findall(r"[a-z0-9]+", value))


def is_top_performer(path: Path) -> bool:
    return "top performer" in normalize(path.stem)


def matching_stem(path: Path) -> str:
    stem = normalize(path.stem)
    stem = re.sub(r"\btop performer\b", " ", stem)
    return " ".join(stem.split())


def score_name(name: str, path: Path) -> float:
    # token_set_ratio handles a PDF such as "First_Last.pdf" when Excel includes a middle name.
    return float(fuzz.token_set_ratio(normalize(name), matching_stem(path)))


def ensure_certificates_extracted() -> None:
    pdfs = list(PDF_DIR.rglob("*.pdf")) if PDF_DIR.exists() else []
    if pdfs:
        return
    if not ZIP_FILE.exists():
        raise FileNotFoundError(f"Cannot find {ZIP_FILE.name} or extracted PDFs.")
    PDF_DIR.mkdir(parents=True, exist_ok=True)
    # Prevent zip-slip/path traversal while extracting.
    with zipfile.ZipFile(ZIP_FILE) as zf:
        root = PDF_DIR.resolve()
        for info in zf.infolist():
            dest = (PDF_DIR / info.filename).resolve()
            if dest != root and root not in dest.parents:
                raise ValueError(f"Unsafe path inside ZIP: {info.filename}")
        zf.extractall(PDF_DIR)


def append_log(student, email, pdf, status, details=""):
    new_file = not LOG_FILE.exists() or LOG_FILE.stat().st_size == 0
    with LOG_FILE.open("a", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=LOG_FIELDS)
        if new_file:
            w.writeheader()
        w.writerow({"timestamp": datetime.now(timezone.utc).isoformat(),
                    "student_name": student, "email": email,
                    "pdf_path": str(pdf.resolve()) if pdf else "",
                    "status": status, "details": details})


def load_log():
    if not LOG_FILE.exists():
        return []
    with LOG_FILE.open("r", newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def already_sent(log, email, pdf):
    e = email.strip().casefold()
    p = str(pdf.resolve()).casefold()
    return any(r.get("status") == "SENT" and r.get("email", "").strip().casefold() == e
               and r.get("pdf_path", "").strip().casefold() == p for r in log)


def best_candidate(name, candidates, used):
    available = [p for p in candidates if p not in used]
    if not available:
        return None, 0, "No unused PDFs available"
    ranked = sorted(((score_name(name, p), p) for p in available), key=lambda x: x[0], reverse=True)
    best_score, best_path = ranked[0]
    second = ranked[1][0] if len(ranked) > 1 else 0
    if best_score < MIN_SCORE:
        return None, best_score, f"Best score {best_score:.1f} below threshold {MIN_SCORE}"
    if len(ranked) > 1 and best_score - second < MIN_GAP:
        return None, best_score, f"Ambiguous: best {best_score:.1f}, second {second:.1f}"
    return best_path, best_score, "Matched"


def build_report(df, pdfs):
    results = []
    eligible_rows = []
    # Detect duplicate records and email collisions before matching.
    seen_pairs = set()
    email_to_rows = {}
    for idx, row in df.iterrows():
        name, email = str(row[NAME_COLUMN]).strip(), str(row[EMAIL_COLUMN]).strip()
        email_to_rows.setdefault(email.casefold(), []).append(idx)
        pair = (normalize(name), email.casefold())
        if pair in seen_pairs:
            results.append({"excel_row": idx + 2, "student_name": name, "email": email,
                            "selected_pdf": "", "match_score": "", "status": "DUPLICATE_ROW",
                            "details": "Duplicate name/email row; not sent twice"})
        else:
            seen_pairs.add(pair)
            eligible_rows.append((idx, row))

    blocked = set()
    for email, indexes in email_to_rows.items():
        if email and len(indexes) > 1:
            names = {normalize(str(df.iloc[i][NAME_COLUMN]).strip()) for i in indexes}
            if len(names) > 1:
                blocked.update(indexes)

    valid_rows = []
    for idx, row in eligible_rows:
        name, email = str(row[NAME_COLUMN]).strip(), str(row[EMAIL_COLUMN]).strip()
        if idx in blocked:
            results.append({"excel_row": idx + 2, "student_name": name, "email": email,
                            "selected_pdf": "", "match_score": "", "status": "DUPLICATE_EMAIL",
                            "details": "Same email assigned to different names; manual review required"})
        elif not name or not normalize(name):
            results.append({"excel_row": idx + 2, "student_name": name, "email": email,
                            "selected_pdf": "", "match_score": "", "status": "MISSING_NAME",
                            "details": "Full name is blank"})
        elif not email or not re.match(r"^[^\s@]+@[^\s@]+\.[^\s@]+$", email):
            results.append({"excel_row": idx + 2, "student_name": name, "email": email,
                            "selected_pdf": "", "match_score": "", "status": "MISSING_OR_INVALID_EMAIL",
                            "details": "Email2 is blank or invalid"})
        else:
            valid_rows.append((idx, name, email))

    special_pdfs = [p for p in pdfs if is_top_performer(p)]
    regular_pdfs = [p for p in pdfs if not is_top_performer(p)]
    used = set()
    assignments = {}

    # First assign Top Performer certificates. If one confidently matches a participant,
    # it replaces that participant's standard certificate.
    for special in special_pdfs:
        candidates = [(score_name(name, special), idx, name, email)
                      for idx, name, email in valid_rows if idx not in assignments]
        candidates.sort(reverse=True, key=lambda x: x[0])
        if not candidates:
            continue
        top_score, idx, name, email = candidates[0]
        second = candidates[1][0] if len(candidates) > 1 else 0
        if top_score >= MIN_SCORE and (len(candidates) == 1 or top_score - second >= MIN_GAP):
            # If a student has more than one special certificate, don't overwrite an assignment.
            if idx not in assignments:
                assignments[idx] = (special, top_score, "MATCHED_TOP_PERFORMER")
                used.add(special)

    for idx, name, email in valid_rows:
        if idx in assignments:
            continue
        path, score, details = best_candidate(name, regular_pdfs, used)
        if path:
            assignments[idx] = (path, score, "MATCHED")
            used.add(path)
        else:
            results.append({"excel_row": idx + 2, "student_name": name, "email": email,
                            "selected_pdf": "", "match_score": round(score, 1),
                            "status": "AMBIGUOUS_OR_NO_MATCH", "details": details})

    for idx, name, email in valid_rows:
        if idx in assignments:
            path, score, status = assignments[idx]
            results.append({"excel_row": idx + 2, "student_name": name, "email": email,
                            "selected_pdf": str(path.resolve()), "match_score": round(score, 1),
                            "status": status,
                            "details": "Top Performer certificate replaces standard certificate" if status == "MATCHED_TOP_PERFORMER" else "Matched; review before sending"})

    for pdf in pdfs:
        if pdf not in used:
            replaced = False
            if not is_top_performer(pdf):
                for idx, (assigned_path, _, assigned_status) in assignments.items():
                    if assigned_status == "MATCHED_TOP_PERFORMER" and score_name(str(df.iloc[idx][NAME_COLUMN]), pdf) >= MIN_SCORE:
                        replaced = True
                        break
            status = ("STANDARD_PDF_REPLACED_BY_TOP_PERFORMER" if replaced else
                      ("UNUSED_STANDARD_PDF" if not is_top_performer(pdf) else "UNMATCHED_TOP_PERFORMER_PDF"))
            details = ("Not sent because the Top Performer certificate is selected for this student" if replaced
                       else "No participant assigned to this PDF")
            results.append({"excel_row": "", "student_name": "", "email": "",
                            "selected_pdf": str(pdf.resolve()), "match_score": "", "status": status,
                            "details": details})

    results.sort(key=lambda r: (int(r["excel_row"]) if str(r["excel_row"]).isdigit() else 9999, r["status"]))
    return results


def send_message(server, recipient, student_name, pdf_path, test=False):
    pdf_path = Path(pdf_path)
    msg = EmailMessage()
    msg["From"] = SENDER
    msg["To"] = recipient
    msg["Subject"] = ("TEST — Your Participation Certificate" if test else "Your Participation Certificate")
    msg.set_content(("This is a TEST email sent to the test address. The attached PDF is the selected certificate for "
                     f"{student_name}. No email was sent to the participant.\n\n" if test else
                     f"Dear {student_name},\n\nThank you for participating in the MLSA Technical Bootcamp. "
                     "Please find your participation certificate attached.\n\nRegards,\nMLSA DYPCET Team\n"))
    msg.add_attachment(pdf_path.read_bytes(), maintype="application", subtype="pdf", filename=pdf_path.name)
    server.send_message(msg)


def write_report(results):
    with REPORT_FILE.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=REPORT_FIELDS, extrasaction="ignore")
        w.writeheader()
        w.writerows(results)


def main():
    ap = argparse.ArgumentParser(description="Match and email existing PDF certificates using Gmail.")
    ap.add_argument("--test-email", help="Send exactly one selected PDF to this test email only.")
    ap.add_argument("--test-name", help="Exact normalized participant name for test; otherwise first matched row.")
    ap.add_argument("--send", action="store_true", help="Send all reviewed matches to participants.")
    args = ap.parse_args()
    if args.test_email and args.send:
        sys.exit("Use either --test-email or --send, not both.")
    if not EXCEL_FILE.exists():
        sys.exit(f"Missing workbook: {EXCEL_FILE}")
    ensure_certificates_extracted()
    df = pd.read_excel(EXCEL_FILE, dtype=str).fillna("")
    missing = [c for c in (NAME_COLUMN, EMAIL_COLUMN) if c not in df.columns]
    if missing:
        sys.exit(f"Required columns missing: {missing}. Found: {list(df.columns)}")
    pdfs = sorted(PDF_DIR.rglob("*.pdf"))
    if not pdfs:
        sys.exit("No PDFs found after extracting certificates.zip")
    results = build_report(df, pdfs)
    log = load_log()

    if args.test_email:
        matches = [r for r in results if r["status"] in ("MATCHED", "MATCHED_TOP_PERFORMER")]
        if args.test_name:
            wanted = normalize(args.test_name)
            matches = [r for r in matches if normalize(r["student_name"]) == wanted]
        else:
            matches = matches[:1]
        if not matches:
            write_report(results)
            sys.exit("No safe matched certificate found for test. Review report.csv")
        if len(matches) != 1:
            write_report(results)
            sys.exit("Test name is not unique. Specify --test-name exactly as shown in report.csv")
        test_recipient = args.test_email.strip()
        if not re.match(r"^[^\s@]+@[^\s@]+\.[^\s@]+$", test_recipient):
            sys.exit("Test email address is invalid.")
        password = getpass.getpass(f"Gmail App Password for {SENDER}: ")
        if not password:
            sys.exit("App Password is required.")
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(GMAIL_HOST, GMAIL_PORT, context=context, timeout=30) as server:
            server.login(SENDER, password)
            send_message(server, test_recipient, matches[0]["student_name"], matches[0]["selected_pdf"], test=True)
        append_log(matches[0]["student_name"], test_recipient, Path(matches[0]["selected_pdf"]),
                   "TEST_SENT", f"Test only; selected participant {matches[0]['student_name']}")
        print(f"Test email submitted to {test_recipient}; participant was not emailed.")

    elif args.send:
        eligible = []
        for r in results:
            if r["status"] not in ("MATCHED", "MATCHED_TOP_PERFORMER"):
                continue
            pdf = Path(r["selected_pdf"])
            if not pdf.is_file():
                r.update(status="MISSING_PDF", details="Selected PDF no longer exists")
            elif already_sent(log, r["email"], pdf):
                r.update(status="ALREADY_SENT", details="Previously logged as SENT")
            else:
                eligible.append(r)
        write_report(results)
        unsafe = [r for r in results if r["status"] in ("AMBIGUOUS_OR_NO_MATCH", "DUPLICATE_EMAIL", "MISSING_NAME", "MISSING_OR_INVALID_EMAIL")]
        print(f"Ready to send: {len(eligible)}; unresolved records: {len(unsafe)}")
        print(f"Review report: {REPORT_FILE}")
        if unsafe:
            print("Safety stop: resolve the unresolved records in report.csv before live sending. Nothing sent.")
            return
        if not eligible:
            print("No eligible unsent certificates. Nothing sent.")
            return
        if input("Type SEND to email these matched participants: ").strip() != "SEND":
            print("Cancelled. No participant emails sent.")
            return
        password = getpass.getpass(f"Gmail App Password for {SENDER}: ")
        if not password:
            sys.exit("App Password is required.")
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(GMAIL_HOST, GMAIL_PORT, context=context, timeout=30) as server:
            server.login(SENDER, password)
            for r in eligible:
                pdf = Path(r["selected_pdf"])
                try:
                    send_message(server, r["email"], r["student_name"], pdf)
                    append_log(r["student_name"], r["email"], pdf, "SENT", "SMTP accepted message")
                    r.update(status="SENT", details="SMTP accepted message")
                    print(f"SENT: {r['student_name']} -> {r['email']}")
                except Exception as exc:
                    append_log(r["student_name"], r["email"], pdf, "ERROR", str(exc))
                    r.update(status="SEND_ERROR", details=str(exc)[:300])
                    print(f"ERROR: {r['student_name']}: {exc}")
                write_report(results)
    else:
        print("DRY RUN ONLY — no emails sent.")
        print(f"PDF files found: {len(pdfs)}")
        for r in results:
            if r["status"] in ("MATCHED", "MATCHED_TOP_PERFORMER"):
                print(f"{r['status']}: {r['student_name']} -> {Path(r['selected_pdf']).name} [{r['match_score']}] -> {r['email']}")
        print("Review report.csv; no email was sent.")

    write_report(results)
    counts = {}
    for r in results:
        counts[r["status"]] = counts.get(r["status"], 0) + 1
    print("\nReport status counts:")
    for k, v in sorted(counts.items()): print(f"  {k}: {v}")
    print(f"Report: {REPORT_FILE}")
    print(f"Send log: {LOG_FILE}")

if __name__ == "__main__":
    main()

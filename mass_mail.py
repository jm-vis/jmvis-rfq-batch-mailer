# mass_mail.py
import os, csv, ssl, time, shutil, smtplib, mimetypes, subprocess, typing
from pathlib import Path
from argparse import ArgumentParser
from email.message import EmailMessage
from email.utils import formataddr, make_msgid
from datetime import date, datetime
from tempfile import TemporaryDirectory

from dotenv import load_dotenv
from docxtpl import DocxTemplate

# MIME types for Excel
mimetypes.add_type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx")
mimetypes.add_type("application/vnd.ms-excel", ".xls")


def load_env() -> dict:
    load_dotenv()
    cfg = {
        "host": os.getenv("SMTP_HOST"),
        "port": int(os.getenv("SMTP_PORT", "587")),
        "user": os.getenv("SMTP_USER"),
        "password": os.getenv("SMTP_PASSWORD"),
        "from_name": os.getenv("FROM_NAME", os.getenv("SMTP_USER", "")),
        "reply_to": os.getenv("REPLY_TO", os.getenv("SMTP_USER", "")),
        "use_ssl": os.getenv("USE_SSL", "false").lower() == "true",
        "sleep": float(os.getenv("SLEEP_SECONDS", "1")),
        "subject_tpl": os.getenv("SUBJECT_TEMPLATE", "RFQ for {company} - documents attached"),
        "deadline": os.getenv("DEADLINE", ""),
        "attach_format": os.getenv("ATTACH_FORMAT", "pdf").lower(),   # pdf or docx
        "max_retries": int(os.getenv("MAX_RETRIES", "3")),
        "body_html_tpl": os.getenv("EMAIL_BODY_HTML_TEMPLATE", ""),    # required
        "logo_path": os.getenv("LOGO_PATH", ""),                       # optional
    }
    missing = [k for k in ("host", "user", "password") if not cfg[k]]
    if missing:
        raise RuntimeError(f"Missing .env variables: {', '.join(missing)}")
    if not cfg["body_html_tpl"]:
        raise RuntimeError("EMAIL_BODY_HTML_TEMPLATE must be set in .env.")
    if cfg["attach_format"] not in {"pdf", "docx"}:
        raise ValueError("ATTACH_FORMAT must be 'pdf' or 'docx'.")
    if cfg["max_retries"] < 1:
        cfg["max_retries"] = 3
    return cfg


def sniff_csv_dialect(path: Path) -> csv.Dialect:
    sample = path.read_text(encoding="utf-8-sig", errors="ignore")[:4096]
    try:
        return csv.Sniffer().sniff(sample, delimiters=";,")
    except Exception:
        class D(csv.Dialect):
            delimiter = ","
            quotechar = '"'
            doublequote = True
            skipinitialspace = False
            lineterminator = "\n"
            quoting = csv.QUOTE_MINIMAL
        return D


def iter_contacts(csv_path: Path):
    dialect = sniff_csv_dialect(csv_path)
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, dialect=dialect)
        required = {"email", "name", "gender", "company"}
        got = {h.strip() for h in (reader.fieldnames or [])}
        if not required.issubset(got):
            raise ValueError(f"contacts.csv must contain columns {sorted(required)}.")
        for row in reader:
            yield {
                "email": row["email"].strip(),
                "name": row["name"].strip(),
                "gender": row["gender"].strip().lower(),
                "company": row["company"].strip(),
            }


def make_salutation(full_name: str, gender: str) -> str:
    parts = [p for p in full_name.split() if p]
    last = parts[-1] if parts else full_name
    if gender == "m":
        return f"Dear Mr {last}"
    if gender == "f":
        return f"Dear Ms {last}"
    return f"Hello {full_name}"


def render_docx(template_path: Path, out_dir: Path, context: dict) -> Path:
    tpl = DocxTemplate(str(template_path))
    tpl.render(context)
    base = f"Cover_Letter_{context['company'].replace(' ', '_')}"
    out_docx = out_dir / f"{base}.docx"
    tpl.save(str(out_docx))
    return out_docx


def convert_docx_to_pdf(docx_path: Path) -> typing.Optional[Path]:
    # Try docx2pdf on Windows/macOS
    try:
        from docx2pdf import convert as d2p_convert  # type: ignore
        out_pdf = docx_path.with_suffix(".pdf")
        d2p_convert(str(docx_path), str(out_pdf))
        if out_pdf.exists():
            return out_pdf
    except Exception:
        pass
    # Fallback LibreOffice
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        outdir = str(docx_path.parent)
        try:
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", "--outdir", outdir, str(docx_path)],
                check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE
            )
            out_pdf = docx_path.with_suffix(".pdf")
            if out_pdf.exists():
                return out_pdf
        except Exception:
            return None
    return None


def add_attachments(msg: EmailMessage, files: typing.Iterable[Path]) -> None:
    for path in files:
        if not path.exists():
            raise FileNotFoundError(f"Attachment not found: {path}")
        ctype, _ = mimetypes.guess_type(str(path))
        if not ctype:
            ctype = "application/octet-stream"
        maintype, subtype = ctype.split("/", 1)
        with path.open("rb") as f:
            data = f.read()
        msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=path.name)


def connect(cfg) -> typing.Tuple[smtplib.SMTP, ssl.SSLContext]:
    context = ssl.create_default_context()
    if cfg["use_ssl"]:
        server = smtplib.SMTP_SSL(cfg["host"], cfg["port"], context=context, timeout=60)
    else:
        server = smtplib.SMTP(cfg["host"], cfg["port"], timeout=60)
        server.ehlo()
        server.starttls(context=context)
        server.ehlo()
    server.login(cfg["user"], cfg["password"])
    return server, context


def render_html_template(path: Path, context: dict) -> str:
    tpl = path.read_text(encoding="utf-8")
    return tpl.format(**context)


def html_to_text(html: str) -> str:
    import re
    text = re.sub(r"<br\s*/?>", "\n", html, flags=re.I)
    text = re.sub(r"</p>", "\n\n", text, flags=re.I)
    text = re.sub(r"<li>", "- ", text, flags=re.I)
    text = re.sub(r"<[^>]+>", "", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def send_one(server: smtplib.SMTP, cfg, to_email: str, subject: str,
             body_text: str, body_html: str, attachments: typing.Iterable[Path],
             logo_path: typing.Optional[Path],
             save_eml_dir: typing.Optional[Path] = None,
             dry_run: bool = False) -> str:
    msg = EmailMessage()
    msg["From"] = formataddr((cfg["from_name"], cfg["user"]))
    msg["To"] = to_email
    msg["Subject"] = subject
    msg["Reply-To"] = cfg["reply_to"]
    msg["Disposition-Notification-To"] = cfg["user"]
    msg["Return-Receipt-To"] = cfg["user"]
    msg["Message-ID"] = make_msgid()

    # multipart/alternative: first text, then html
    msg.set_content(body_text)
    # replace {logo_cid} with real CID before adding html
    html_for_send = body_html
    img_cid = None
    if logo_path:
        img_cid = make_msgid()[1:-1]  # strip <>
        html_for_send = html_for_send.replace("{logo_cid}", f"cid:{img_cid}")
    msg.add_alternative(html_for_send, subtype="html")

    # embed logo inline into the HTML part
    if logo_path and logo_path.exists():
        ctype, _ = mimetypes.guess_type(str(logo_path))
        if not ctype:
            ctype = "image/png"
        maintype, subtype = ctype.split("/", 1)
        with open(logo_path, "rb") as f:
            # HTML part is the second payload (index 1)
            msg.get_payload()[1].add_related(
                f.read(), maintype=maintype, subtype=subtype, cid=img_cid, filename=logo_path.name
            )

    add_attachments(msg, attachments)

    if save_eml_dir:
        save_eml_dir.mkdir(parents=True, exist_ok=True)
        eml = save_eml_dir / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{to_email.replace('@','_at_')}.eml"
        eml.write_bytes(msg.as_bytes())

    if dry_run:
        return msg["Message-ID"]

    server.send_message(msg)
    return msg["Message-ID"]


def resolve_path(arg_path: str) -> Path:
    """
    Robust: versuche arg wie gegeben, dann relativ zu CWD, dann relativ zum Skriptordner.
    """
    p = Path(arg_path)
    if p.exists():
        return p.resolve()
    for base in (Path.cwd(), Path(__file__).parent):
        q = base / arg_path
        if q.exists():
            return q.resolve()
    return p.resolve()  # existiert nicht, aber für Fehlermeldung brauchbar


def main():
    p = ArgumentParser(description="RFQ batch mailer with HTML template, inline logo and CSV logging")
    p.add_argument("--contacts", help="Path to contacts.csv")
    p.add_argument("--docx", required=True, help="Path to cover_letter_template.docx")
    p.add_argument("--xlsx", required=True, help="Path to specifications.xlsx")
    p.add_argument("--limit", type=int, default=0, help="Send only first N")
    p.add_argument("--dry-run", action="store_true", help="Preview only, do not send")
    p.add_argument("--retry-from-log", help="Path to previous send_log_*.csv to retry failures")
    p.add_argument("--save-eml-out", help="Directory to write .eml messages for testing")
    p.add_argument("--write-status-csv", help="Write derived contacts status CSV to this path")
    args = p.parse_args()

    cfg = load_env()

    # Pfade robust auflösen und prüfen
    docx_tpl = resolve_path(args.docx)
    xlsx_path = resolve_path(args.xlsx)
    if not docx_tpl.exists():
        raise FileNotFoundError(f"Cover letter template not found: {docx_tpl}")
    if not xlsx_path.exists():
        raise FileNotFoundError(f"Excel not found: {xlsx_path}")

    html_tpl_path = resolve_path(cfg["body_html_tpl"])
    if not html_tpl_path.exists():
        raise FileNotFoundError(f"HTML template not found: {html_tpl_path}")

    logo_path = resolve_path(cfg["logo_path"]) if cfg["logo_path"] else None
    if cfg["logo_path"] and not Path(cfg["logo_path"]).exists():
        # Hinweis, kein Abbruch
        print(f"Logo not found at {cfg['logo_path']}. Sending without logo.")

    # Log vorbereiten
    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = Path(f"send_log_{run_id}.csv")
    log_fields = ["run_id","timestamp","attempt","email","name","gender","company",
                  "subject","attachments","status","message_id","error"]
    log_file = log_path.open("w", newline="", encoding="utf-8")
    log = csv.DictWriter(log_file, fieldnames=log_fields)
    log.writeheader()

    # Kontakte laden
    contacts = []
    prev_attempts = {}
    last_error_map = {}

    if args.retry_from_log:
        src = resolve_path(args.retry_from_log)
        if not src.exists():
            raise FileNotFoundError(f"Retry log not found: {src}")
        with src.open("r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            tmp = {}
            for row in reader:
                email = (row.get("email") or "").strip()
                if not email:
                    continue
                status = (row.get("status") or "").strip().lower()
                if email not in tmp:
                    tmp[email] = {"email": email, "name": row.get("name","").strip(),
                                  "gender": (row.get("gender") or "x").strip().lower(),
                                  "company": row.get("company","").strip(),
                                  "prev": 0, "sent": False, "err": ""}
                if status == "sent":
                    tmp[email]["sent"] = True
                elif status in {"failed","aborted_max_retries","skipped_max_retries"}:
                    tmp[email]["prev"] += 1
                    if row.get("error"):
                        tmp[email]["err"] = row["error"]
            for e, d in tmp.items():
                if not d["sent"] and d["prev"] < cfg["max_retries"]:
                    contacts.append({"email": d["email"], "name": d["name"], "gender": d["gender"], "company": d["company"]})
                    prev_attempts[d["email"]] = d["prev"]
                    last_error_map[d["email"]] = d["err"]
    else:
        if not args.contacts:
            raise ValueError("Without --retry-from-log you must pass --contacts.")
        contacts = list(iter_contacts(resolve_path(args.contacts)))

    if args.limit > 0:
        contacts = contacts[:args.limit]

    # Dry-run Vorschau
    if args.dry_run:
        with TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            for c in contacts[:3]:
                sal = make_salutation(c["name"], c["gender"])
                subject = cfg["subject_tpl"].format(company=c["company"])
                ctx = {
                    "salutation": sal,
                    "company": c["company"],
                    "deadline": cfg["deadline"],
                    "from_name": cfg["from_name"],
                    "reply_to": cfg["reply_to"],
                    "today": date.today().strftime("%m/%d/%Y"),
                    "logo_cid": "{logo_cid}",
                }
                out_docx = render_docx(docx_tpl, tmpdir, ctx)
                letter = out_docx
                if cfg["attach_format"] == "pdf":
                    out_pdf = convert_docx_to_pdf(out_docx)
                    letter = out_pdf if out_pdf else out_docx
                html_body = render_html_template(html_tpl_path, ctx)
                text_body = html_to_text(html_body)
                print("Preview:")
                print("To:", c["email"])
                print("Subject:", subject)
                print(text_body)
                print("Cover letter:", letter.name)
                print("Excel:", xlsx_path.name)
                print("-" * 40)
        log_file.close()
        print(f"Log saved: {log_path}")
        return

    # Versand
    server, _ = connect(cfg)
    ok = fail = skipped = 0
    latest_status = {}

    with TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        for c in contacts:
            email = c["email"]
            already = prev_attempts.get(email, 0)
            if already >= cfg["max_retries"]:
                skipped += 1
                subj = cfg["subject_tpl"].format(company=c["company"])
                log.writerow({
                    "run_id": run_id, "timestamp": datetime.now().isoformat(timespec="seconds"),
                    "attempt": already, "email": email, "name": c["name"], "gender": c["gender"], "company": c["company"],
                    "subject": subj, "attachments": "", "status": "skipped_max_retries", "message_id": "", "error": last_error_map.get(email,"")
                })
                latest_status[email] = {"status": "skipped_max_retries", "attempt": already, "message_id": "", "error": last_error_map.get(email,"")}
                print(f"SKIP {email} due to {already} previous failures")
                continue

            attempt = already + 1
            while attempt <= cfg["max_retries"]:
                try:
                    sal = make_salutation(c["name"], c["gender"])
                    subject = cfg["subject_tpl"].format(company=c["company"])
                    ctx = {
                        "salutation": sal,
                        "company": c["company"],
                        "deadline": cfg["deadline"],
                        "from_name": cfg["from_name"],
                        "reply_to": cfg["reply_to"],
                        "today": date.today().strftime("%m/%d/%Y"),
                        "logo_cid": "{logo_cid}",
                    }
                    out_docx = render_docx(docx_tpl, tmpdir, ctx)
                    letter = out_docx
                    if cfg["attach_format"] == "pdf":
                        out_pdf = convert_docx_to_pdf(out_docx)
                        letter = out_pdf if out_pdf else out_docx

                    attachments = [letter, xlsx_path]
                    html_body = render_html_template(html_tpl_path, ctx)
                    text_body = html_to_text(html_body)

                    msg_id = send_one(
                        server, cfg, email, subject, text_body, html_body, attachments,
                        logo_path=logo_path if (logo_path and logo_path.exists()) else None,
                        save_eml_dir=Path(args.save_eml_out) if args.save_eml_out else None,
                        dry_run=False
                    )
                    ok += 1
                    log.writerow({
                        "run_id": run_id, "timestamp": datetime.now().isoformat(timespec="seconds"),
                        "attempt": attempt, "email": email, "name": c["name"], "gender": c["gender"], "company": c["company"],
                        "subject": subject, "attachments": ", ".join([p.name for p in attachments]),
                        "status": "sent", "message_id": msg_id, "error": ""
                    })
                    latest_status[email] = {"status": "sent", "attempt": attempt, "message_id": msg_id, "error": ""}
                    print(f"OK {email} [{msg_id}]")
                    time.sleep(cfg["sleep"])
                    break
                except Exception as e:
                    fail += 1
                    err = str(e)
                    subject = cfg["subject_tpl"].format(company=c["company"])
                    log.writerow({
                        "run_id": run_id, "timestamp": datetime.now().isoformat(timespec="seconds"),
                        "attempt": attempt, "email": email, "name": c["name"], "gender": c["gender"], "company": c["company"],
                        "subject": subject, "attachments": "", "status": "failed",
                        "message_id": "", "error": err
                    })
                    latest_status[email] = {"status": "failed", "attempt": attempt, "message_id": "", "error": err}
                    print(f"FAIL attempt {attempt}: {email} -> {err}")
                    if attempt >= cfg["max_retries"]:
                        log.writerow({
                            "run_id": run_id, "timestamp": datetime.now().isoformat(timespec="seconds"),
                            "attempt": attempt, "email": email, "name": c["name"], "gender": c["gender"], "company": c["company"],
                            "subject": subject, "attachments": "", "status": "aborted_max_retries",
                            "message_id": "", "error": err
                        })
                        latest_status[email] = {"status": "aborted_max_retries", "attempt": attempt, "message_id": "", "error": err}
                        print(f"ABORT {email} after {attempt} failed attempts")
                        break
                    time.sleep(cfg["sleep"])
                    attempt += 1

    server.quit()
    log_file.close()
    print(f"Done. Sent: {ok}, Failed: {fail}, Skipped: {skipped}")
    print(f"Log saved: {log_path}")

    if args.write_status_csv:
        outp = resolve_path(args.write_status_csv)
        with outp.open("w", newline="", encoding="utf-8") as f:
            fieldnames = ["email","name","gender","company","last_status","attempt","message_id","error"]
            w = csv.DictWriter(f, fieldnames=fieldnames); w.writeheader()
            contact_map = {c["email"]: c for c in contacts}
            for email, st in latest_status.items():
                c = contact_map.get(email, {"name":"","gender":"","company":""})
                w.writerow({
                    "email": email,
                    "name": c.get("name",""),
                    "gender": c.get("gender",""),
                    "company": c.get("company",""),
                    "last_status": st["status"],
                    "attempt": st["attempt"],
                    "message_id": st["message_id"],
                    "error": st["error"],
                })
        print(f"Status CSV written: {outp}")


if __name__ == "__main__":
    main()

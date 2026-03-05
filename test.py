import imaplib
import email
from email.policy import default

imap = imaplib.IMAP4_SSL("imap.mail.me.com", 993)
imap.login("cb1.10.10.10@icloud.com", "xxxx")
imap.select("heizoel")

_, msg_ids = imap.search(None, "ALL")
ids = msg_ids[0].split()
print(f"Gefundene Mails: {len(ids)}\n")

for mid in ids:
    _, data = imap.fetch(mid, "(BODY[])")
    
    raw = None
    for part in data:
        if isinstance(part, tuple) and len(part) >= 2 and isinstance(part[1], bytes):
            raw = part[1]
            break

    if not raw:
        print(f"Mail {mid}: Kein Inhalt")
        continue

    msg = email.message_from_bytes(raw, policy=default)

    print("=" * 50)
    print("Betreff:", msg["subject"])
    print("Von:    ", msg["from"])
    print("An:     ", msg["to"])
    print("Datum:  ", msg["date"])

    # Body extrahieren
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                body = part.get_content()
                break
    else:
        body = msg.get_content()

    print("Body:\n", body[:500])  # Erste 500 Zeichen

imap.logout()
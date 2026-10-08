#!/usr/bin/env python3
"""Build the password-locked invoice site.

    python tools/lock_site.py SRC_DIR OUT_DIR

SRC_DIR holds the plain page (index.html) and the signature/stamp images.
It must NOT be committed: only OUT_DIR (encrypted page + non-secret files)
goes into the public repository. The password is read from the
SITE_USER / SITE_PASSWORD environment variables or asked for interactively.
"""
import base64, getpass, json, os, shutil, sys
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes

ITERATIONS = 600_000
SECRET_IMAGES = ["sig-prepared.png", "sig-authorized.png", "stamp.png"]
PUBLIC_FILES = ["logo.png", "Carlito-Regular.ttf", "Carlito-Bold.ttf"]

def b64(b): return base64.b64encode(b).decode()

def main(src, out):
    user = (os.environ.get("SITE_USER") or input("Username: ")).strip().lower()
    pw = os.environ.get("SITE_PASSWORD") or getpass.getpass("Password: ")
    if not user or len(pw) < 8: sys.exit("Use a username and a password of at least 8 characters.")
    html = open(os.path.join(src, "index.html"), encoding="utf-8").read()
    for name in SECRET_IMAGES:  # signatures travel inside the encrypted page only
        data = open(os.path.join(src, name), "rb").read()
        needle = f'file: "{name}"'
        if needle not in html: sys.exit(f"{needle} not found in index.html")
        html = html.replace(needle, f'file: "data:image/png;base64,{b64(data)}"')
    salt, iv = os.urandom(16), os.urandom(12)
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITERATIONS).derive((user + "\n" + pw).encode())
    ct = AESGCM(key).encrypt(iv, html.encode("utf-8"), None)
    payload = json.dumps({"v": 1, "iter": ITERATIONS, "salt": b64(salt), "iv": b64(iv), "data": b64(ct)})
    gate = open(os.path.join(os.path.dirname(__file__), "login.html"), encoding="utf-8").read()
    os.makedirs(out, exist_ok=True)
    open(os.path.join(out, "index.html"), "w", encoding="utf-8").write(gate.replace("__PAYLOAD__", payload))
    for name in PUBLIC_FILES:
        for d in (src, out):
            p = os.path.join(d, name)
            if os.path.exists(p) and d == src: shutil.copy(p, os.path.join(out, name))
    for name in SECRET_IMAGES:
        p = os.path.join(out, name)
        if os.path.exists(p): os.remove(p)
    print(f"Locked site written to {out}")

if __name__ == "__main__":
    if len(sys.argv) != 3: sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2])

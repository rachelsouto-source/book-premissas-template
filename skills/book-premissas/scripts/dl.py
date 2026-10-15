"""Baixa arquivos do Drive (via Composio) e reduz imagens a 2400 px.

  python dl.py PASTA_SAIDA ID[=nome] [ID[=nome] ...]
  python dl.py PASTA_SAIDA @lista.txt          # uma linha por arquivo: ID=nome (ou so ID)

Imagens (.jpg/.png/.jpeg/.webp) sao salvas como JPEG com lado maior <= 2400 px.
"""
import io
import os
import sys
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import composio  # noqa: E402

MAX = 2400


def fetch(fid):
    r = composio.run("GOOGLEDRIVE_DOWNLOAD_FILE", {"file_id": fid})
    d = r.get("data", r)
    c = d.get("downloaded_file_content") or d.get("file") or {}
    url = c.get("s3url") or c.get("url")
    if not url:
        raise RuntimeError(f"sem url para {fid}: {str(r)[:300]}")
    return c.get("name") or fid, urllib.request.urlopen(url, timeout=600).read()


def save(out_dir, name, data):
    ext = os.path.splitext(name)[1].lower()
    path = os.path.join(out_dir, name)
    if ext in (".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".bmp"):
        from PIL import Image
        im = Image.open(io.BytesIO(data))
        im = im.convert("RGB")
        w, h = im.size
        s = MAX / max(w, h)
        if s < 1:
            im = im.resize((int(w * s), int(h * s)), Image.LANCZOS)
        path = os.path.splitext(path)[0] + ".jpg"
        im.save(path, "JPEG", quality=90)
    else:
        open(path, "wb").write(data)
    return path


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    out_dir = sys.argv[1]
    os.makedirs(out_dir, exist_ok=True)
    items = []
    for a in sys.argv[2:]:
        if a.startswith("@"):
            items += [ln.strip() for ln in open(a[1:], encoding="utf-8") if ln.strip() and not ln.startswith("#")]
        else:
            items.append(a)
    import time
    for it in items:
        fid, _, name = it.partition("=")
        if name and (os.path.exists(os.path.join(out_dir, name + ".jpg")) or os.path.exists(os.path.join(out_dir, name))):
            print("ja existe", name)
            continue
        data = None
        for tent in range(6):
            try:
                orig, data = fetch(fid)
                break
            except Exception as e:  # noqa: BLE001
                msg = str(e)
                if "Quota" in msg or "quota" in msg or "429" in msg:
                    print("quota; aguardando 30 s", fid)
                    time.sleep(30)
                    continue
                print("ERRO", fid, msg[:200])
                break
        if data is None:
            print("ERRO (desistiu)", fid)
            continue
        if name:
            ext = os.path.splitext(orig)[1] or ".jpg"
            name = name if os.path.splitext(name)[1] else name + ext
        else:
            name = orig
        print(save(out_dir, name, data), len(data))


if __name__ == "__main__":
    main()

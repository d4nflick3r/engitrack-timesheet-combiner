import os
import sys
import glob

PWA_BLOCK = """\
    <link rel="manifest" href="/app/static/manifest.json" />
    <meta name="theme-color" content="#1F4E79" />
    <meta name="mobile-web-app-capable" content="yes" />
    <meta name="apple-mobile-web-app-capable" content="yes" />
    <meta name="apple-mobile-web-app-title" content="EngiTrack" />
    <meta name="application-name" content="EngiTrack" />
    <link rel="apple-touch-icon" href="/app/static/icons/icon-192.png" />
"""

ANCHOR = '    <link rel="shortcut icon" href="./favicon.png" />'


def find_index_html():
    import streamlit
    st_dir = os.path.dirname(streamlit.__file__)
    candidate = os.path.join(st_dir, "static", "index.html")
    if os.path.isfile(candidate):
        return candidate
    patterns = [
        "/home/runner/workspace/.pythonlibs/lib/python*/site-packages/streamlit/static/index.html",
        "/nix/store/*/site-packages/streamlit/static/index.html",
    ]
    for p in patterns:
        matches = glob.glob(p)
        if matches:
            return matches[0]
    return None


def patch():
    if getattr(sys, "frozen", False):
        return
    path = find_index_html()
    if not path:
        return
    try:
        with open(path, "r") as f:
            content = f.read()
        if 'rel="manifest"' in content:
            return
        patched = content.replace(ANCHOR, ANCHOR + "\n" + PWA_BLOCK, 1)
        if patched == content:
            return
        with open(path, "w") as f:
            f.write(patched)
    except OSError:
        pass


if __name__ == "__main__":
    patch()
    print("done")

from PIL import Image

ico = Image.open("byf.ico")
sizes = []

try:
    for i in range(getattr(ico, "n_frames", 1)):
        ico.seek(i)
        sizes.append(ico.size)
except Exception:
    sizes.append(ico.size)

print("format:", ico.format)
print("n_frames:", getattr(ico, "n_frames", 1))
print("sizes:", sorted(set(sizes)))
print("info sizes:", ico.info.get("sizes"))
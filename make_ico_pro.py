from PIL import Image

# Cambiá esto si tu archivo tiene otro nombre
INPUT_PNG = "logo-1024.png"
OUTPUT_ICO = "logo.ico"

# Tamaños profesionales multi-resolución para Windows
SIZES = [
    (16, 16),
    (32, 32),
    (48, 48),
    (64, 64),
    (128, 128),
    (256, 256),
]

def main():
    img = Image.open(INPUT_PNG).convert("RGBA")

    # Opcional: si querés forzar fondo transparente
    # img = img.convert("RGBA")

    img.save(
        OUTPUT_ICO,
        format="ICO",
        sizes=SIZES
    )

    print(f"Icono profesional creado correctamente: {OUTPUT_ICO}")

if __name__ == "__main__":
    main()
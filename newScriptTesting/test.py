import cv2
from PIL import Image
from pylibdmtx.pylibdmtx import decode

for i in range(1, 5):
    print(f"\n=== ssd{i}.png ===")
    img = cv2.imread(f"ssd{i}.png", cv2.IMREAD_GRAYSCALE)
    up = cv2.resize(img, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)

    results = decode(Image.fromarray(up), timeout=5000)
    for r in results:
        print(f"  Data: {r.data.decode()}")

    h, w = up.shape
    crop = up[:, :w // 3]
    results = decode(Image.fromarray(crop), timeout=5000)
    for r in results:
        print(f"  Cropped - Data: {r.data.decode()}")

    if not results:
        print("  Nothing found.")
"""
SSD Serial Number Reader
Uses a webcam + OCR to detect and extract serial numbers from SSD stickers.
Serial numbers typically follow labels like: SN:, S/N:, S/T:, Serial No., etc.
 
Requirements:
    pip install opencv-python pytesseract Pillow numpy
 
Also install Tesseract OCR engine:
    - Windows: https://github.com/UB-Mannheim/tesseract/wiki
    - macOS:   brew install tesseract
    - Linux:   sudo apt install tesseract-ocr
"""
 
import re
import sys
import time
import threading
from pathlib import Path
 
import cv2
import numpy as np
import pytesseract
from PIL import Image
 
# ── Configuration ─────────────────────────────────────────────────────────────
 
# On Windows you may need to point to the Tesseract executable, e.g.:
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
 
# Patterns for serial number labels (case-insensitive)
SN_LABEL_PATTERNS = [
    r"S[/\-]?[NT][:.]\s*([A-Z0-9\-]{4,30})",   # SN: / S/N: / ST: / S-N: / S-T:
    r"Serial\s*No\.?\s*[:\-]?\s*([A-Z0-9\-]{4,30})",
    r"Serial\s*#\s*[:\-]?\s*([A-Z0-9\-]{4,30})",
    r"SER\.?\s*NO\.?\s*[:\-]?\s*([A-Z0-9\-]{4,30})",
    r"S\.No\.?\s*[:\-]?\s*([A-Z0-9\-]{4,30})",
]
 
COMPILED_PATTERNS = [re.compile(p, re.IGNORECASE) for p in SN_LABEL_PATTERNS]
 
CAMERA_INDEX  = 0         # Change if your webcam is on a different index
FRAME_SKIP    = 3         # Process every Nth frame (reduces CPU load)
CONFIRM_COUNT = 3         # How many consecutive matches before announcing
WINDOW_NAME   = "SSD Serial Number Reader  |  Press Q to quit, S to save"
 
# ── Image pre-processing helpers ──────────────────────────────────────────────
 
def preprocess(frame: np.ndarray) -> list[np.ndarray]:
    """
    Return several preprocessed variants of the frame so OCR can try each.
    More variants = better coverage for different lighting / print conditions.
    """
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    variants = []
 
    # 1. Upscaled grayscale (small text reads much better at 2×)
    upscaled = cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    variants.append(upscaled)
 
    # 2. Adaptive threshold (handles uneven lighting / shadows)
    thresh = cv2.adaptiveThreshold(
        upscaled, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY, 31, 10
    )
    variants.append(thresh)
 
    # 3. Inverted threshold (for white-on-dark stickers)
    variants.append(cv2.bitwise_not(thresh))
 
    # 4. CLAHE-enhanced (improves low-contrast stickers)
    clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
    enhanced = clahe.apply(upscaled)
    variants.append(enhanced)
 
    return variants
 
 
def extract_text(image: np.ndarray) -> str:
    """Run Tesseract OCR on a pre-processed image and return the raw text."""
    pil_img = Image.fromarray(image)
    # PSM 6 = assume a single uniform block of text (good for labels)
    config = "--oem 3 --psm 6 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789/:.-_ "
    return pytesseract.image_to_string(pil_img, config=config)
 
 
# ── Serial number extraction ───────────────────────────────────────────────────
 
def find_serial_numbers(text: str) -> list[str]:
    """Search the OCR text for known serial-number label patterns."""
    found = []
    for pattern in COMPILED_PATTERNS:
        for match in pattern.finditer(text):
            sn = match.group(1).strip().upper()
            if sn not in found:
                found.append(sn)
    return found
 
 
# ── Overlay drawing ───────────────────────────────────────────────────────────
 
def draw_overlay(frame: np.ndarray, serial_numbers: list[str], fps: float) -> np.ndarray:
    overlay = frame.copy()
    h, w = overlay.shape[:2]
 
    # Semi-transparent top bar
    cv2.rectangle(overlay, (0, 0), (w, 50), (20, 20, 20), -1)
    cv2.addWeighted(overlay, 0.75, frame, 0.25, 0, overlay)
 
    cv2.putText(overlay, f"FPS: {fps:.1f}", (10, 34),
                cv2.FONT_HERSHEY_SIMPLEX, 0.7, (180, 180, 180), 1, cv2.LINE_AA)
    cv2.putText(overlay, "Point camera at SSD sticker", (w // 2 - 160, 34),
                cv2.FONT_HERSHEY_SIMPLEX, 0.7, (200, 200, 200), 1, cv2.LINE_AA)
 
    # ROI guide rectangle
    roi_x1, roi_y1 = w // 8, h // 4
    roi_x2, roi_y2 = 7 * w // 8, 3 * h // 4
    cv2.rectangle(overlay, (roi_x1, roi_y1), (roi_x2, roi_y2), (0, 255, 128), 2)
    cv2.putText(overlay, "Align sticker here", (roi_x1 + 8, roi_y1 - 8),
                cv2.FONT_HERSHEY_SIMPLEX, 0.55, (0, 255, 128), 1, cv2.LINE_AA)
 
    # Serial number results
    if serial_numbers:
        box_y = roi_y2 + 16
        for i, sn in enumerate(serial_numbers):
            label = f"Serial: {sn}"
            (tw, th), _ = cv2.getTextSize(label, cv2.FONT_HERSHEY_DUPLEX, 0.9, 2)
            bx1, by1 = roi_x1, box_y + i * (th + 20)
            cv2.rectangle(overlay, (bx1 - 6, by1 - th - 6), (bx1 + tw + 6, by1 + 6),
                          (0, 60, 0), -1)
            cv2.rectangle(overlay, (bx1 - 6, by1 - th - 6), (bx1 + tw + 6, by1 + 6),
                          (0, 220, 80), 2)
            cv2.putText(overlay, label, (bx1, by1),
                        cv2.FONT_HERSHEY_DUPLEX, 0.9, (0, 255, 80), 2, cv2.LINE_AA)
    else:
        cv2.putText(overlay, "No serial number detected", (roi_x1 + 8, roi_y2 + 30),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.65, (80, 80, 255), 1, cv2.LINE_AA)
 
    return overlay
 
 
# ── OCR worker (runs in a background thread) ──────────────────────────────────
 
class OCRWorker:
    def __init__(self):
        self.latest_frame: np.ndarray | None = None
        self.results: list[str] = []
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
 
    def submit(self, frame: np.ndarray):
        with self._lock:
            self.latest_frame = frame.copy()
 
    def get_results(self) -> list[str]:
        with self._lock:
            return list(self.results)
 
    def stop(self):
        self._stop.set()
 
    def _run(self):
        while not self._stop.is_set():
            frame = None
            with self._lock:
                if self.latest_frame is not None:
                    frame = self.latest_frame
                    self.latest_frame = None   # consume it
 
            if frame is None:
                time.sleep(0.05)
                continue
 
            h, w = frame.shape[:2]
            # Crop to the guide ROI before OCR
            roi = frame[h // 4: 3 * h // 4, w // 8: 7 * w // 8]
 
            all_serials: list[str] = []
            for variant in preprocess(roi):
                text = extract_text(variant)
                all_serials.extend(find_serial_numbers(text))
 
            # Deduplicate while preserving order
            seen: set[str] = set()
            unique = []
            for s in all_serials:
                if s not in seen:
                    seen.add(s)
                    unique.append(s)
 
            with self._lock:
                self.results = unique
 
 
# ── Main loop ─────────────────────────────────────────────────────────────────
 
def main():
    cap = cv2.VideoCapture(CAMERA_INDEX)
    if not cap.isOpened():
        print(f"[ERROR] Cannot open camera index {CAMERA_INDEX}.")
        sys.exit(1)
 
    # Request a decent resolution
    cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
 
    worker = OCRWorker()
    frame_idx = 0
    t_prev = time.time()
    fps = 0.0
 
    # Confirmation state: only announce a SN after CONFIRM_COUNT consecutive reads
    confirm_buffer: dict[str, int] = {}
    confirmed_serials: list[str] = []
 
    print("=" * 60)
    print("  SSD Serial Number Reader")
    print("  Hold your SSD in front of the camera.")
    print("  Press  Q  to quit,  S  to save a snapshot.")
    print("=" * 60)
 
    save_dir = Path("ssd_snapshots")
 
    while True:
        ret, frame = cap.read()
        if not ret:
            print("[WARN] Frame capture failed — retrying…")
            time.sleep(0.1)
            continue
 
        frame_idx += 1
 
        # Submit every Nth frame to the OCR worker
        if frame_idx % FRAME_SKIP == 0:
            worker.submit(frame)
 
        # Update FPS
        now = time.time()
        fps = 0.9 * fps + 0.1 / max(now - t_prev, 1e-6)
        t_prev = now
 
        # Retrieve latest OCR results
        raw_results = worker.get_results()
 
        # Confirmation logic
        new_counts: dict[str, int] = {}
        for sn in raw_results:
            new_counts[sn] = confirm_buffer.get(sn, 0) + 1
        confirm_buffer = new_counts
 
        confirmed_serials = [
            sn for sn, cnt in confirm_buffer.items() if cnt >= CONFIRM_COUNT
        ]
 
        if confirmed_serials:
            for sn in confirmed_serials:
                if sn not in getattr(main, "_announced", set()):
                    if not hasattr(main, "_announced"):
                        main._announced = set()
                    main._announced.add(sn)
                    print(f"[DETECTED] Serial Number: {sn}")
 
        # Draw and show
        display = draw_overlay(frame, confirmed_serials, fps)
        cv2.imshow(WINDOW_NAME, display)
 
        key = cv2.waitKey(1) & 0xFF
        if key == ord("q") or key == 27:          # Q or Esc → quit
            break
        elif key == ord("s"):                      # S → save snapshot
            save_dir.mkdir(exist_ok=True)
            ts = time.strftime("%Y%m%d_%H%M%S")
            path = save_dir / f"ssd_{ts}.jpg"
            cv2.imwrite(str(path), display)
            print(f"[SAVED] Snapshot saved to {path}")
 
    worker.stop()
    cap.release()
    cv2.destroyAllWindows()
    print("\nBye!")
 
 
if __name__ == "__main__":
    main()
 
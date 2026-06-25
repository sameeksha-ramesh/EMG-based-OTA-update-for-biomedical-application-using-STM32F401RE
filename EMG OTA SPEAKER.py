"""
EMG Bootloader — Python Serial Monitor + Voice Feedback
========================================================
Listens on the STM32 UART (115200 baud) and speaks
"OTA update done, application is now running"
when it detects the [OTA]  OTA_UPDATE_DONE sentinel line.

Requirements:
    pip install pyserial

Voice uses the built-in Windows Speech API (win32com).
On Linux/Mac, replace with pyttsx3 (also provided as fallback).

Usage:
    python emg_ota_listener.py          # auto-detect port
    python emg_ota_listener.py COM5     # specify port
    python emg_ota_listener.py /dev/ttyACM0
"""

import sys
import time
import threading
import serial
import serial.tools.list_ports

# ── Voice backend ─────────────────────────────────────────────────────────────
def make_speaker():
    """Return a speak(text) callable using the best available TTS."""
    # Try Windows SAPI first
    try:
        import win32com.client
        sapi = win32com.client.Dispatch("SAPI.SpVoice")
        def speak_sapi(text):
            sapi.Speak(text)
        print("[VOICE] Using Windows Speech API (SAPI)")
        return speak_sapi
    except Exception:
        pass

    # Fallback: pyttsx3 (cross-platform)
    try:
        import pyttsx3
        engine = pyttsx3.init()
        engine.setProperty("rate", 160)
        def speak_pyttsx3(text):
            engine.say(text)
            engine.runAndWait()
        print("[VOICE] Using pyttsx3")
        return speak_pyttsx3
    except Exception:
        pass

    # Last resort: print only
    print("[VOICE] No TTS available — printing only")
    def speak_print(text):
        print(f"[VOICE] >> {text}")
    return speak_print


# ── Port detection ────────────────────────────────────────────────────────────
def find_stm32_port():
    """Heuristically find the ST-Link virtual COM port."""
    candidates = list(serial.tools.list_ports.comports())
    # Prefer known STM32 / ST-Link vendor strings
    for p in candidates:
        desc = (p.description or "").lower()
        if any(k in desc for k in ["stm32", "st-link", "nucleo", "stlink"]):
            return p.device
    # Fall back to first available port
    if candidates:
        return candidates[0].device
    return None


# ── Serial monitor ────────────────────────────────────────────────────────────
OTA_DONE_SENTINEL   = "OTA_UPDATE_DONE"
BAUD                = 115200

# Colour codes (ANSI) — disabled on Windows if not supported
CYAN   = "\033[96m"
GREEN  = "\033[92m"
YELLOW = "\033[93m"
RED    = "\033[91m"
RESET  = "\033[0m"

def coloured(color, text):
    try:
        return f"{color}{text}{RESET}"
    except Exception:
        return text


def monitor(port: str, speak):
    print(coloured(CYAN, f"\n[MONITOR] Connecting to {port} @ {BAUD} baud …"))
    try:
        ser = serial.Serial(port, BAUD, timeout=1)
    except serial.SerialException as e:
        print(coloured(RED, f"[ERROR] Cannot open {port}: {e}"))
        sys.exit(1)

    print(coloured(GREEN, f"[MONITOR] Connected. Waiting for STM32 output …\n"))

    ota_announced = False

    try:
        while True:
            raw = ser.readline()
            if not raw:
                continue

            try:
                line = raw.decode("utf-8", errors="replace").rstrip()
            except Exception:
                continue

            # Timestamp prefix
            ts = time.strftime("%H:%M:%S")
            print(f"[{ts}] {line}")

            # ── Detect EMG trigger ─────────────────────────────────────────
            if "EMG confirmed" in line:
                print(coloured(YELLOW, "         ↑ EMG threshold hit — OTA handover triggered"))

            # ── Detect OTA completion ──────────────────────────────────────
            if OTA_DONE_SENTINEL in line and not ota_announced:
                ota_announced = True
                msg = "O T A update done. Application is now running."
                print(coloured(GREEN, f"\n{'='*55}"))
                print(coloured(GREEN, f"  ✓  {msg}"))
                print(coloured(GREEN, f"{'='*55}\n"))

                # Speak in a background thread so serial reading continues
                threading.Thread(target=speak, args=(msg,), daemon=True).start()

            # ── Detect app running (reset flag so next OTA is announced) ───
            if "[APP]" in line and ota_announced:
                ota_announced = False   # ready for next cycle

    except KeyboardInterrupt:
        print(coloured(CYAN, "\n[MONITOR] Stopped by user."))
    finally:
        ser.close()


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    speak = make_speaker()

    if len(sys.argv) > 1:
        port = sys.argv[1]
    else:
        port = find_stm32_port()
        if port is None:
            print(coloured(RED, "[ERROR] No serial port found. Pass port as argument."))
            print("        Usage: python emg_ota_listener.py COM5")
            sys.exit(1)
        print(coloured(CYAN, f"[AUTO]  Detected port: {port}"))

    monitor(port, speak)

import os
import sys
import threading
import winsound

def get_sound_path(sound_filename: str) -> str:
    """Locate sound file in source dir or PyInstaller sys._MEIPASS dir."""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    snd_path = os.path.join(base_dir, "sounds", sound_filename)
    if os.path.exists(snd_path):
        return snd_path

    if hasattr(sys, "_MEIPASS"):
        meipass_snd = os.path.join(sys._MEIPASS, "sounds", sound_filename)
        if os.path.exists(meipass_snd):
            return meipass_snd

    return ""

def _play_sound_async(sound_filename: str):
    path = get_sound_path(sound_filename)
    if path and os.path.exists(path):
        try:
            winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_ASYNC)
        except Exception as e:
            print(f"[SoundManager] PlaySound error: {e}")

def play_click():
    threading.Thread(target=_play_sound_async, args=("click.wav",), daemon=True).start()

def play_xp():
    threading.Thread(target=_play_sound_async, args=("xp_gain.wav",), daemon=True).start()

def play_victory():
    threading.Thread(target=_play_sound_async, args=("victory.wav",), daemon=True).start()

if __name__ == "__main__":
    print("Testing sound manager...")
    play_click()

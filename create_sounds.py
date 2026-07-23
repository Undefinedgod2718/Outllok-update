import os
import wave
import math
import struct

def generate_wav(filename, duration, freq_func, volume=0.5, sample_rate=44100):
    num_samples = int(duration * sample_rate)
    os.makedirs(os.path.dirname(filename), exist_ok=True)
    
    with wave.open(filename, 'w') as wav_file:
        wav_file.setnchannels(1)     # Mono
        wav_file.setsampwidth(2)    # 16-bit
        wav_file.setframerate(sample_rate)
        
        for i in range(num_samples):
            t = float(i) / sample_rate
            freq = freq_func(t, duration)
            # Square / Triangle retro 8-bit wave synth
            val = math.sin(2.0 * math.pi * freq * t)
            if val > 0:
                sample_val = int(32767 * volume)
            else:
                sample_val = int(-32767 * volume)
                
            # Envelope decay
            decay = max(0.0, 1.0 - (t / duration))
            sample_val = int(sample_val * decay)
            
            data = struct.pack('<h', sample_val)
            wav_file.writeframesraw(data)

def create_all_sounds():
    sounds_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sounds")
    os.makedirs(sounds_dir, exist_ok=True)

    # 1. click.wav (Short frequency sweep 800Hz -> 400Hz)
    click_path = os.path.join(sounds_dir, "click.wav")
    generate_wav(
        click_path,
        duration=0.06,
        freq_func=lambda t, d: 900.0 - (500.0 * (t / d)),
        volume=0.4
    )

    # 2. xp_gain.wav (Rising pitch chime 1000Hz -> 1600Hz)
    xp_path = os.path.join(sounds_dir, "xp_gain.wav")
    generate_wav(
        xp_path,
        duration=0.12,
        freq_func=lambda t, d: 1000.0 + (600.0 * (t / d)),
        volume=0.5
    )

    # 3. victory.wav (4-Note retro fanfare C5:523Hz, E5:659Hz, G5:784Hz, C6:1046Hz)
    victory_path = os.path.join(sounds_dir, "victory.wav")
    def victory_freq(t, d):
        if t < 0.12:
            return 523.25 # C5
        elif t < 0.24:
            return 659.25 # E5
        elif t < 0.36:
            return 783.99 # G5
        else:
            return 1046.50 # C6

    generate_wav(
        victory_path,
        duration=0.60,
        freq_func=victory_freq,
        volume=0.5
    )

    print(f"✅ 所有 8-bit Retro 音效檔已成功生成於 {sounds_dir}:")
    print(f"   - {click_path}")
    print(f"   - {xp_path}")
    print(f"   - {victory_path}")

if __name__ == "__main__":
    create_all_sounds()

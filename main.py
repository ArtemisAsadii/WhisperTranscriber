import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import whisper
from docx import Document
import time
import os
import urllib.request

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

stop_flag = False
model = None  



def browse_file():
    file_path = filedialog.askopenfilename(
        filetypes=(("Audio files", "*.mp3 *.wav"), ("All files", "*.*"))
    )
    if file_path:
        entry_file.set(file_path)

def save_to_word(text):
    save_path = filedialog.asksaveasfilename(
        defaultextension=".docx",
        filetypes=[("Word Document", "*.docx")]
    )
    if save_path:
        try:
            doc = Document()
            doc.add_paragraph(text)
            doc.save(save_path)
            messagebox.showinfo("Success", f"Text saved to {save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Word file: {e}")

def download_model_thread(model_name="base"):
    global model
    global stop_flag
    stop_flag = False

    try:

        for i in range(1, 101):
            if stop_flag:
                download_progress.set(0)
                return
            download_progress.set(i/100)
            root.update_idletasks()
            time.sleep(0.03)
        model = whisper.load_model(model_name)
        download_progress.set(1.0)
        messagebox.showinfo("Model Downloaded", f"Model '{model_name}' is ready.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to download model: {e}")

def transcribe_audio_thread():
    global stop_flag
    stop_flag = False
    audio_path = entry_file.get()
    if not audio_path:
        messagebox.showwarning("Warning", "Please select an audio file")
        return
    if model is None:
        messagebox.showwarning("Warning", "Please download the model first")
        return

    try:
        result = model.transcribe(audio_path, verbose=False)
        # شبیه‌سازی درصد تبدیل
        for i in range(1, 101):
            if stop_flag:
                convert_progress.set(0)
                return
            convert_progress.set(i/100)
            root.update_idletasks()
            time.sleep(0.02)
        convert_progress.set(1.0)
        save_to_word(result["text"])
    except Exception as e:
        messagebox.showerror("Error", f"Failed to transcribe audio: {e}")
    finally:
        convert_progress.set(0)

def start_download_model():
    threading.Thread(target=download_model_thread, daemon=True).start()

def start_transcription():
    threading.Thread(target=transcribe_audio_thread, daemon=True).start()

def cancel_all():
    global stop_flag
    stop_flag = True

# ======================= GUI =======================

root = ctk.CTk()
root.title("Whisper Audio to Text Converter")
root.geometry("650x400")

ctk.CTkLabel(root, text="Download Whisper Model:", font=("Arial", 14)).pack(pady=10)
ctk.CTkButton(root, text="Download Model", command=start_download_model, width=200, fg_color="#2196F3").pack(pady=5)
download_progress = ctk.CTkProgressBar(root, width=500)
download_progress.set(0)
download_progress.pack(pady=10)

ctk.CTkLabel(root, text="Select Audio File (MP3/WAV):", font=("Arial", 14)).pack(pady=10)
entry_file = ctk.StringVar()
ctk.CTkEntry(root, textvariable=entry_file, width=450).pack(pady=5)
ctk.CTkButton(root, text="Browse", command=browse_file, width=120, fg_color="#4CAF50").pack(pady=5)

ctk.CTkButton(root, text="Transcribe & Save to Word", command=start_transcription, width=250, fg_color="#2196F3").pack(pady=10)
convert_progress = ctk.CTkProgressBar(root, width=500)
convert_progress.set(0)
convert_progress.pack(pady=10)

ctk.CTkButton(root, text="Cancel", command=cancel_all, width=250, fg_color="#f44336").pack(pady=10)

root.mainloop()

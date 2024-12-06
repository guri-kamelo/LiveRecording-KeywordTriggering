import speech_recognition as sr
import threading
import queue

class AudioToTextTranscriber:
    def __init__(self):
        self.recognizer = sr.Recognizer()
        self.microphone = None
        self.is_recording = False
        self.audio_queue = queue.Queue()

    def get_microphone_devices(self):
        return sr.Microphone.list_microphone_names()

    def start_transcription(self, device_index, callback):
        """Starts transcription on the selected device.

        Args:
            device_index (int): Index of the microphone to use.
            callback (callable): Function to handle recognized text.
        """
        self.is_recording = True
        self.microphone = sr.Microphone(device_index=device_index)

        threading.Thread(target=self.capture_audio, daemon=True).start()
        threading.Thread(target=self.process_audio, args=(callback,), daemon=True).start()

    def stop_transcription(self):
        self.is_recording = False

    def capture_audio(self):
        with self.microphone as source:
            self.recognizer.adjust_for_ambient_noise(source)
            while self.is_recording:
                try:
                    audio = self.recognizer.listen(source, timeout=1, phrase_time_limit=2)
                    self.audio_queue.put(audio)
                except sr.WaitTimeoutError:
                    continue

    def process_audio(self, callback):
        while self.is_recording or not self.audio_queue.empty():
            try:
                audio = self.audio_queue.get(timeout=1)
                text = self.recognizer.recognize_google(audio, language="he-IL")
                callback(f" {text} ")
            except sr.UnknownValueError:
                callback("-")
            except sr.RequestError as e:
                callback(f"API error: {e}")
                break
            except queue.Empty:
                pass

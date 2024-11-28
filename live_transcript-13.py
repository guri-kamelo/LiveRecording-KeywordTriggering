import speech_recognition as sr
import time
import difflib
import os
import threading
import tkinter as tk
from tkinter import scrolledtext, messagebox
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import csv
import pandas as pd
import comtypes.client
import pythoncom
import time
import pyaudio  # Import the PyAudio module
from pydub import AudioSegment
from pydub.silence import detect_nonsilent
import io

# Initialize recognizer
recognizer = sr.Recognizer()

# Directory to save audio files and the text file
script_dir = os.path.dirname(os.path.abspath(__file__))
AUDIO_SAVE_PATH = os.path.join(script_dir, "AudioTrigger")
TEXT_FILE_PATH = os.path.join(AUDIO_SAVE_PATH, "transcribed_texts.txt")
os.makedirs(AUDIO_SAVE_PATH, exist_ok=True)  # Create directory if it doesn't exist

# Global flag to control recording and text updates
is_recording = False
flat_text={}

flat_text['text'] = ""
flat_text['num_of_lines'] = 0
# Global variable to store transcribed text
collected_texts = []
start_index = "1.0"  # Start searching from the beginning of the text
current_highlight_color = "red"  # Start with red
# Global variable to store the selected microphone device index
selected_device_index = 1  # Default device index
# Global variables
is_recording = False
trigger_pointer = 1  # Pointer to track which trigger we are waiting for
collected_Arrow = False
trigger_keywords = []  # List to hold all keywords from the CSV
trigger_listbox = []
recording_thread = None  # To track the recording thread
lock = threading.Lock()  # To ensure thread safety

def open_ppt(ppt_file):
    """
    Opens a PowerPoint file and returns the PowerPoint application and presentation object.
    """
    # Initialize PowerPoint application
   
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1

    # Open the PowerPoint presentation
    presentation = powerpoint.Presentations.Open(ppt_file)

    return powerpoint, presentation



def play_slide_with_animations(powerpoint, presentation, slide_number):
    """
    Moves to the specific slide and displays it with animations.
    """
    # Set the presentation to slide show mode from the specific slide
    slide_show = presentation.SlideShowSettings
    slide_show.StartingSlide = slide_number
    slide_show.EndingSlide = slide_number
    slide_show.AdvanceMode = 1  # Automatic mode, allowing animations to proceed
    slide_show.Run()
    # Access the slide show window to control the view
    slide_show_window = powerpoint.SlideShowWindows(1)
    # Move to the specific slide if not already there
    slide_show_window.View.GotoSlide(slide_number)

# Function to read keyword triggers from a CSV file
def load_triggers_from_csv(csv_file):
    global trigger_keywords
    df = pd.read_excel(csv_file)
    trigger_keywords = df["Keywords"].dropna().tolist()

    print(f"Loaded triggers: {trigger_keywords}")

def increase_trigger_position(trigger_listbox):
    """
    Moves the trigger_pointer to the next trigger (down) and updates the highlight.
    """
    global trigger_pointer
    global collected_Arrow
    if trigger_pointer <= len(trigger_keywords):  # Ensure pointer stays within bounds
        trigger_pointer += 1
        update_trigger_position()
        highlight_current_trigger(trigger_listbox, trigger_pointer)
        collected_Arrow = True

def decrease_trigger_position(trigger_listbox):
    """
    Moves the trigger_pointer to the previous trigger (up) and updates the highlight.
    """
    global trigger_pointer
    global collected_Arrow

    if trigger_pointer > 1:  # Ensure pointer stays within bounds
        trigger_pointer -= 1
        dec_trigger_position()
        highlight_current_trigger(trigger_listbox, trigger_pointer)
        collected_Arrow = True

def highlight_current_trigger(trigger_listbox, trigger_pointer):
    """
    Highlights the current trigger in the listbox based on the trigger pointer.
    
    Args:
    trigger_listbox: The listbox containing trigger keywords.
    trigger_pointer: The index of the current trigger to be highlighted.
    """
    # Remove previous highlights
    for i in range(trigger_listbox.size()):
        trigger_listbox.itemconfig(i, {'bg': 'white'})  # Reset background to white

    # Highlight the current trigger
    trigger_listbox.itemconfig(trigger_pointer-1, {'bg': 'yellow'})  # Set the background to yellow
    trigger_listbox.see(trigger_pointer-1)  # Scroll to the highlighted item

# Wrapper function to monitor keyword triggers during the show
# Modify full_show_triggers to call `update_trigger_position`
def full_show_triggers(csv_file):
    global trigger_pointer, trigger_keywords
    load_triggers_from_csv(csv_file)  # Load triggers from the CSV file
    
    if not trigger_keywords:
        print("No triggers found in the CSV file.")
        return

    def monitor_triggers():
        global trigger_pointer
        global collected_Arrow
        global flat_text
        global start_index
        pythoncom.CoInitialize()  # Initialize COM in this thread
        
        ppt_file_path = os.path.join(script_dir, "C:\\Users\\gurik\\OneDrive\\Documents\\מצא את המטמון הרצל.pptx")
        powerpoint, presentation = open_ppt(ppt_file_path)

        while trigger_pointer <= len(trigger_keywords):
            if collected_texts:
                flat_text['text'] += " " + collected_texts[0]
                flat_text['num_of_lines'] += 1
                print(f"Flat Text: {flat_text}")
                current_keyword = trigger_keywords[trigger_pointer - 1]  # Get the current trigger keyword
                # Call the keyword trigger function for the current keyword
                if keyword_trigger(current_keyword, powerpoint, presentation,flat_text):
                    flat_text['text'] = ""
                    flat_text['num_of_lines'] = 0
                    trigger_pointer += 1  # Move to the next trigger in the list
                    # Highlight the current trigger in the listbox
                    start_index = increment_index(start_index, line_increment=1, column_increment=0)
                    # Step 2: Move to the specific slide and display it
                    play_slide_with_animations(powerpoint, presentation, trigger_pointer+1)
                    highlight_current_trigger(trigger_listbox, trigger_pointer)
                if collected_texts and collected_texts[0] != "":
                    collected_texts.pop(0)

            elif collected_Arrow:
                play_slide_with_animations(powerpoint, presentation, trigger_pointer)
                collected_Arrow = False

        

            time.sleep(1)
        pythoncom.CoUninitialize() 
    # Start monitoring for keyword triggers in a separate thread
    threading.Thread(target=monitor_triggers, daemon=True).start()


def record_audio(output_dir=AUDIO_SAVE_PATH, debug=False, min_duration_ms=5000 ):
    """
    Continuously records audio and saves segments. Each segment will be at least min_duration_ms long.
    
    Args:
        output_dir (str): Directory to save audio chunks.
        debug (bool): Flag to save recorded audio regardless of silence detection.
        min_duration_ms (int): Minimum duration for each audio segment in milliseconds.
        device_index (int): Index of the audio device to use (None to use default).
    """
    global selected_device_index
    global trigger_keywords
    global is_recording
    current_keyword = trigger_keywords[trigger_pointer - 1]
    keyword_length = len(current_keyword.split())
    min_duration_ms = (keyword_length+2)*1000
    device_index=selected_device_index
    # Audio settings
    RATE = 44100  # Sampling rate
    CHUNK = 1024*8  # Audio chunk size
    FORMAT = pyaudio.paInt16  # Audio format
    CHANNELS = 1  # Mono recording

    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # PyAudio setup
    p = pyaudio.PyAudio()

    # Open audio stream with optional device index
    stream = p.open(format=FORMAT, channels=CHANNELS, rate=RATE, input=True,
                    input_device_index=device_index, frames_per_buffer=CHUNK)

    print("Recording started. Press Ctrl+C to stop.")

    try:
        audio_buffer = b''  # Use bytes to accumulate audio data
        count = 0
        min_chunks = (min_duration_ms * RATE) // (CHUNK * 1000)  # Calculate the number of chunks for min_duration_ms

        while is_recording:
            # Read audio data from the microphone
            data = stream.read(CHUNK)
            audio_buffer += data

            # If the buffer is large enough (at least min_duration_ms long), save the segment
            if len(audio_buffer) >= min_chunks * CHUNK * p.get_sample_size(FORMAT):
                # Write the audio buffer to a BytesIO object for pydub to process
                audio_io = io.BytesIO(audio_buffer)
                
                # Create an AudioSegment from the raw audio data
                audio_segment = AudioSegment.from_raw(audio_io, sample_width=p.get_sample_size(FORMAT),
                                                      frame_rate=RATE, channels=CHANNELS,read_ahead_limit=1024*4)

                # Save the audio segment to a file
                audio_filename = os.path.join(output_dir, f"audio_debug_segment_{count}.wav")
                audio_segment.export(audio_filename, format="wav")
                #transcribe_audio(audio_filename)
                print(f"Saved debug audio segment: {audio_filename}")
                count += 1
                silence_duration=2000
                silence_thresh=-80
                #nonsilent_ranges = detect_nonsilent(audio_buffer, min_silence_len=silence_duration, silence_thresh=silence_thresh)
                nonsilent_ranges = detect_nonsilent(audio_segment, min_silence_len=silence_duration, silence_thresh=silence_thresh)
                print(f"nonsilent_ranges{nonsilent_ranges}") 
                # Reset the audio buffer after saving to avoid growing memory usage
                audio_buffer = b''

    except KeyboardInterrupt:
        print("Recording stopped.")
    finally:
        # Close PyAudio stream
        stream.stop_stream()
        stream.close()
        p.terminate()

        print("Audio recording finished.")

# Watchdog event handler to monitor new WAV files and trigger transcription
class AudioFileHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.src_path.endswith(".wav"):
            print(f"New file detected: {event.src_path}")
            #transcribe_audio(event.src_path)
            transcription_thread = threading.Thread(target=transcribe_audio, args=(event.src_path,))
            transcription_thread.start()

# Function to transcribe a given audio file
def transcribe_audio(audio_file):
    global collected_texts
    debug = 1

    # Load the audio file
    with sr.AudioFile(audio_file) as source:
        audio = recognizer.record(source)

    # Try to recognize the speech in the audio
    try:
        text = recognizer.recognize_google(audio, language="he-IL")
        print(f"Transcribed text from {audio_file}: {text}")
        collected_texts.append(text)

        # Append the transcribed text to the file
        with open(TEXT_FILE_PATH, "a", encoding="utf-8") as file:
            file.write(text + "\n")

        # Update the live transcript display
        update_transcript(text)

        # Remove the audio file after transcription
        if not debug :
            os.remove(audio_file)
            print(f"Audio file {audio_file} removed successfully.")

    except sr.UnknownValueError:

        print(f"Could not understand the audio in {audio_file}")
        if not debug :
            os.remove(audio_file)
            print(f"Audio file {audio_file} removed successfully.")

    except sr.RequestError as e:
        print(f"Could not request results from Google Speech Recognition service; {e}")

# Function to search for a keyword in the live transcript
def keyword_trigger(keyword, powerpoint, presentation,flat_text):
    global collected_texts
    global start_index

    global trigger_pointer

    triggered_keywords = []  # List to store the triggered keywords

    keyword_words = keyword.split()  # Split the keyword sentence into words
    keyword_length = len(keyword_words)  # Get the number of words in the keyword
    trigger_threshold = min(4, keyword_length // 2)  # Set the threshold to half of the keyword words

    if flat_text:
        text = flat_text['text'].lower()  # Take the first transcribed sentence
        sentence_words = text.split()  # Split the transcribed sentence into words
        
        match_count = 0
        
        # Compare each word in the sentence with each word in the keyword
        for keyword_word in keyword_words:
            for sentence_word in sentence_words:
                if difflib.SequenceMatcher(None, keyword_word, sentence_word).ratio() > 0.8:
                    match_count += 1
                    triggered_keywords.append(keyword_word)  # Add the keyword word to the list
                    break  # Move to the next keyword word after finding a match
        
        # Check if the number of matching words exceeds the threshold
        if match_count > trigger_threshold:
            print(f"Found keyword trigger |{triggered_keywords}| in text: {text}")
            print(f"Match count                  : {match_count}/{trigger_threshold}")
            highlight_keywords(text, ' '.join(triggered_keywords))


            return triggered_keywords  # Return the list of triggered keywords
        else:
            print(f"Did not find keyword trigger in text: {text}")
            print(f"Match count                  : {match_count}/{trigger_threshold}")
            start_index = increment_index(start_index, line_increment=1, column_increment=0)

            return []  # Return an empty list if the threshold isn't met

    return []  # Return an empty list if there is no text to process


# Function to update the live transcript window
def update_transcript(text):
    # Align the text to the right to support Hebrew
    transcript_window.tag_configure("right", justify='right')
    transcript_window.insert(tk.END, text + "\n", "right")
    transcript_window.see(tk.END)

# New sub-window to show trigger position and completion status
def update_trigger_position():
    global trigger_pointer
    if trigger_pointer < len(trigger_keywords):
        trigger_status_window.config(text=f"Trigger Position: {trigger_pointer}/{len(trigger_keywords)}")
    else:
        trigger_status_window.config(text="Done")
        stop_recording()  # Stop recording when all triggers are processed

def dec_trigger_position():
    global trigger_pointer
    trigger_status_window.config(text=f"Trigger Position: {trigger_pointer}/{len(trigger_keywords)}")


# Function to highlight each keyword in the transcript with unique tags
def increment_index(start_index, line_increment=0, column_increment=0):
    # Split the start_index into line and column
    line, column = map(int, start_index.split('.'))
    
    # Increment the line and column
    line += line_increment
    column += column_increment
    
    # Ensure the column does not go below zero
    if column < 0 or line_increment > 0:
        column = 0
    
    # Return the new index in "line.column" format
    print(f"Updated counter position from: {start_index} to {line}.{column}")
    return f"{line}.{column}"

# Function to highlight only triggered words
def highlight_keywords(text, keyword):
    print(f"highlight_keywords |{keyword}| in text: {text}")

    keyword_list = keyword.split()  # Split the keyword into individual words
    
    global start_index
    global current_highlight_color

    new_color = "green" if current_highlight_color == "red" else "red"
    current_highlight_color = new_color  # Alternate between red and green
    end_index = start_index
    # Loop through each word in the keyword list
    for keyword_word in keyword_list:
        # Normalize the keyword for case-insensitive matching
        keyword_lower = keyword_word.lower()

        # Search and highlight the word within the transcribed text
        while True:
            tmp_start_index = transcript_window.search(keyword_lower, start_index, tk.END, nocase=True)
            
            # If no match is found, stop searching for this word
            if tmp_start_index != '' :
                start_index=tmp_start_index
                # Calculate the end index
                end_index = increment_index(start_index, line_increment=0, column_increment=len(keyword_word))
            
                # Apply the highlight color
                transcript_window.tag_add(f"highlight_{keyword_word}", start_index, end_index)
                transcript_window.tag_config(f"highlight_{keyword_word}", background=new_color)

                # Move the column search index forward
                start_index = increment_index(start_index, line_increment=0, column_increment=len(keyword_word))

                break
            else :
                break


    # Update the trigger position display
    update_trigger_position()

# Function to clean the audio directory
def clean_audio_directory():
    # Get a list of all files in the audio directory
    files = os.listdir(AUDIO_SAVE_PATH)

    # Iterate through each file in the directory
    for file in files:
        # Check if the file is a WAV file
        if file.endswith(".wav"):
            file_path = os.path.join(AUDIO_SAVE_PATH, file)
            try:
                # Remove the file
                os.remove(file_path)
                print(f"Deleted old audio file: {file_path}")
            except Exception as e:
                print(f"Error deleting file {file_path}: {e}")


# Start recording button action
def start_recording():
    global is_recording, recording_thread

    with lock:
        if is_recording:
            print("Already recording!")
            return

        clean_audio_directory()  # Clean the directory before starting the recording
        is_recording = True

    recording_thread = threading.Thread(target=record_audio)
    recording_thread.start()
    print("Recording started.")

# Stop recording button action
def stop_recording():
    global is_recording, recording_thread

    with lock:
        if not is_recording:
            print("Recording is not active.")
            return

        is_recording = False  # Set flag to stop the recording
        print("Stopping recording...")

    if recording_thread:
        #recording_thread.join()  # Wait for the thread to finish
        recording_thread = None  # Reset the thread reference

    print("Recording stopped.")

# Function to monitor the directory for new WAV files
def monitor_audio_directory():
    event_handler = AudioFileHandler()
    observer = Observer()
    observer.schedule(event_handler, AUDIO_SAVE_PATH, recursive=False)
    observer.start()
    print("Started monitoring audio directory for new files.")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

def add_trigger_listbox(root, csv_file_path):
    """
    Reads trigger keywords from a CSV file and adds a Listbox next to the transcript window.
    
    Args:
    root: The Tkinter root window.
    csv_file_path: The path to the CSV file containing the trigger keywords.
    """
    # Create a Listbox widget
    #trigger_listbox = tk.Listbox(root, height=20, width=40)
    trigger_listbox = tk.Listbox(root, height=20, width=40, justify='right', font=("Arial", 12))

    trigger_listbox.pack(side=tk.RIGHT, padx=10, pady=10, fill=tk.Y)

    # Read triggers from the CSV file and insert them into the listbox
    try:
        df = pd.read_excel(csv_file_path)

        # הנחה שהעמודה בשם "Keywords" מכילה את הטריגרים
        for index, trigger in enumerate(df["Keywords"].dropna()):  # הסרת NaN
            trigger_listbox.insert(tk.END, f"{index + 1}. {trigger}")
    except FileNotFoundError:
        trigger_listbox.insert(tk.END, "Error: CSV file not found.")
    except Exception as e:
        trigger_listbox.insert(tk.END, f"Error reading CSV file: {str(e)}")

    return trigger_listbox

# Function to open a popup for selecting the recording device
def select_device():
    def apply_selection():
        global selected_device_index
        selected_device_index = device_var.get()  # Update the selected device index
        print(f"Selected recording device index: {selected_device_index}")
        device_window.destroy()

    # Get the list of available microphone devices
    devices = sr.Microphone.list_microphone_names()

    if not devices:
        messagebox.showerror("No Input Devices", "No input devices (microphones) found.")
        return

    # Create a new popup window
    device_window = tk.Toplevel(root)
    device_window.title("Select Recording Device")

    # Variable to store the selected device index
    device_var = tk.IntVar(value=selected_device_index)

    # Display available devices with radio buttons
    for i, device_name in enumerate(devices):
        tk.Radiobutton(device_window, text=device_name, variable=device_var, value=i).pack(anchor=tk.W)

    # Apply button to confirm the selection
    apply_button = tk.Button(device_window, text="Apply", command=apply_selection)
    apply_button.pack(pady=10)

# GUI setup
ppt_file_path = os.path.join(script_dir, "tmunot_ofaa.pptx")

root = tk.Tk()
root.title("Speech Recognition with Keyword Highlighting")

# Transcript display window
transcript_window = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=80, height=20)
transcript_window.pack(side=tk.LEFT, padx=10, pady=10, fill=tk.BOTH, expand=True)

# Start and stop buttons
start_button = tk.Button(root, text="Start Recording", command=start_recording)
start_button.pack(padx=5, pady=5, side=tk.LEFT)

stop_button = tk.Button(root, text="Stop Recording", command=stop_recording)
stop_button.pack(padx=5, pady=5, side=tk.RIGHT)

# Device selection button
device_button = tk.Button(root, text="Select Recording Device", command=select_device)
device_button.pack(padx=5, pady=5, side=tk.BOTTOM)

# Example: Run full_show_triggers in a separate thread
csv_file_path = os.path.join(script_dir, "show_triggers.xlsx")  # Example CSV file path

# Create a thread for the full_show_triggers function and start it
show_trigger_thread = threading.Thread(target=full_show_triggers, args=(csv_file_path,), daemon=True)
show_trigger_thread.start()

# Sub-window to display current trigger status
trigger_status_window = tk.Label(root, text="Trigger Position: 0/0", font=("Arial", 12))
trigger_status_window.pack(padx=10, pady=5)

# Monitor the directory for new WAV files in a separate thread
threading.Thread(target=monitor_audio_directory, daemon=True).start()

# Add the listbox for the trigger keywords using the CSV file
trigger_listbox=add_trigger_listbox(root, csv_file_path)

control_frame = tk.Frame(root)
control_frame.pack(pady=5)

up_button = tk.Button(control_frame, text="↑", command=lambda: decrease_trigger_position(trigger_listbox))
up_button.pack(side=tk.LEFT, padx=5)

down_button = tk.Button(control_frame, text="↓", command=lambda: increase_trigger_position(trigger_listbox))
down_button.pack(side=tk.LEFT, padx=5)

# Run the GUI
root.mainloop()

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

# Initialize recognizer
recognizer = sr.Recognizer()

# Directory to save audio files and the text file
script_dir = os.path.dirname(os.path.abspath(__file__))
AUDIO_SAVE_PATH = os.path.join(script_dir, "AudioTrigger")
TEXT_FILE_PATH = os.path.join(AUDIO_SAVE_PATH, "transcribed_texts.txt")
os.makedirs(AUDIO_SAVE_PATH, exist_ok=True)  # Create directory if it doesn't exist

# Global flag to control recording and text updates
is_recording = False

# Global variable to store transcribed text
collected_texts = []
start_index = "1.0"  # Start searching from the beginning of the text
current_highlight_color = "red"  # Start with red
# Global variable to store the selected microphone device index
selected_device_index = 1  # Default device index
# Global variables
is_recording = False
collected_texts = []
trigger_pointer = 1  # Pointer to track which trigger we are waiting for
collected_Arrow = False
trigger_keywords = []  # List to hold all keywords from the CSV
trigger_listbox = []


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


import time

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

        pythoncom.CoInitialize()  # Initialize COM in this thread
        
        ppt_file_path = os.path.join(script_dir, "tmunot_ofaa.pptx")
        powerpoint, presentation = open_ppt(ppt_file_path)

        while trigger_pointer <= min(len(trigger_keywords),3):
            if collected_texts:
                current_keyword = trigger_keywords[trigger_pointer - 1]  # Get the current trigger keyword
                # Call the keyword trigger function for the current keyword
                if keyword_trigger(current_keyword, powerpoint, presentation):
                    trigger_pointer += 1  # Move to the next trigger in the list
                    # Highlight the current trigger in the listbox
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

# Function to record audio continuously and save as WAV files every 10 seconds
def record_audio(duration=10):
    global is_recording, selected_device_index
    mic = sr.Microphone(device_index=selected_device_index)
    count = 0  # To keep track of how many audio segments have been saved
    
    with mic as source:
        recognizer.adjust_for_ambient_noise(source)  # Adjust for noise only once
        print("Listening for audio...")

        while is_recording:
            print(f"Recording for the next {duration} seconds...")

            # Record the audio for the specified duration
            audio = recognizer.listen(source, timeout=0, phrase_time_limit=duration)

            # Save audio as a WAV file
            audio_filename = os.path.join(AUDIO_SAVE_PATH, f"audio_segment_{count}.wav")
            with open(audio_filename, "wb") as f:
                f.write(audio.get_wav_data())

            print(f"Audio saved as {audio_filename}")
            count += 1

            time.sleep(0.1)  # Small delay to ensure audio files don't overlap

# Watchdog event handler to monitor new WAV files and trigger transcription
class AudioFileHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.src_path.endswith(".wav"):
            print(f"New file detected: {event.src_path}")
            transcribe_audio(event.src_path)

# Function to transcribe a given audio file
def transcribe_audio(audio_file):
    global collected_texts

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
        os.remove(audio_file)
        print(f"Audio file {audio_file} removed successfully.")

    except sr.UnknownValueError:
        print(f"Could not understand the audio in {audio_file}")
    except sr.RequestError as e:
        print(f"Could not request results from Google Speech Recognition service; {e}")

# Function to search for a keyword in the live transcript
def keyword_trigger(keyword, powerpoint, presentation):
    global collected_texts
    global start_index

    global trigger_pointer

    triggered_keywords = []  # List to store the triggered keywords

    keyword_words = keyword.split()  # Split the keyword sentence into words
    keyword_length = len(keyword_words)  # Get the number of words in the keyword
    trigger_threshold = max(1, keyword_length // 2)  # Set the threshold to half of the keyword words

    if collected_texts:
        text = collected_texts[0]  # Take the first transcribed sentence
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
            start_index = increment_index(start_index, line_increment=1, column_increment=0)
            # Step 2: Move to the specific slide and display it
            play_slide_with_animations(powerpoint, presentation, trigger_pointer+1)
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
    global is_recording
    clean_audio_directory()  # Clean the directory before starting the recording
    is_recording = True
    threading.Thread(target=record_audio).start()

# Stop recording button action
def stop_recording():
    global is_recording
    is_recording = False
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

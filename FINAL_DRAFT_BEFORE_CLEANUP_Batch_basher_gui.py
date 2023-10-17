import logging
import subprocess
import os
import csv
import re
import base64
import io
#import webbrowser
import threading
import time
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
#import ttkthemes
from PIL import Image, ImageTk

# Configure and setup environment vars
logging.basicConfig(filename='batch_lip_basher_log.txt', level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s') # Maybe log when the program starts so we catch all exceptions kay
dialog_export_file_path = None # Initialize the variable, important otherwise code throws a hissy fit

# Define a dictionary to store the loaded images
loaded_images = {}

def load_base64_image(image_data):
    # Decode the Base64 image data
    image_bytes = base64.b64decode(image_data)

    # Load the image from bytes
    img = Image.open(io.BytesIO(image_bytes))

    return img

def close_splash_screen():
    # Destroy the splash screen window
    splash_screen.destroy()
    time.sleep(0)
    window.deiconify()
    window.destroy()
    # Show the existing main window

# Create the main window
window = tk.Tk()
window.title("Initializing, Please Wait...")

# Load the splash screen image
base64_image = "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"
splash_image = load_base64_image(base64_image)

# Create the splash screen window as a child of the main window
splash_screen = tk.Toplevel(window)
splash_screen.title("Splash Screen")

# Hide the title bar and window buttons
splash_screen.overrideredirect(True)

# Display the splash screen image
splash_photo = ImageTk.PhotoImage(splash_image)
splash_label = tk.Label(splash_screen, image=splash_photo)
splash_label.pack()

# Center the splash screen on the screen
screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()
splash_width = splash_screen.winfo_reqwidth()
splash_height = splash_screen.winfo_reqheight()
x = (screen_width - splash_width) // 2
y = (screen_height - splash_height) // 2
splash_screen.geometry(f"{splash_width}x{splash_height}+{x}+{y}")

# Hide the existing main window initially
window.withdraw()

# Update the main window after a delay (e.g., 3 seconds)
splash_screen.after(1700, close_splash_screen)
splash_screen.mainloop()

# Start the Tkinter event loop for the main window
window.mainloop()

def process_files():
    game = game_entry.get()
    language = language_entry.get()
    fonix_data = "FonixData.cdf"  # Set default value for Fonix Data
    output_folder = output_entry.get()
    total_files = len(treeview.get_children())  # Total number of files

    # Retrieve audio files and voice lines from the table
    audio_files = []
    voice_lines = []
    for item in treeview.get_children():
        audio_files.append(treeview.item(item, "values")[0])
        voice_lines.append(treeview.item(item, "values")[1])

    if len(audio_files) != len(voice_lines):
        messagebox.showerror("Error", "Number of audio files and voice lines should match.")
        return

    # Disable the "Process Files" button during processing
    process_button.config(state=tk.DISABLED)

    # Update the loading bar in the GUI from the main thread
    def update_loading_bar(completed_files):
        progress = (completed_files / total_files) * 100
        loading_bar.config(value=progress)

    # Process files in a separate thread
    def process_files_thread():
        completed_files = 0
        for index, audio_file in enumerate(audio_files):
            if save_location_checkbox.get():
                audio_path = os.path.dirname(audio_file)
                lip_filename = os.path.splitext(os.path.basename(audio_file))[0] + ".lip"
                lip_file = os.path.join(audio_path, lip_filename)
            else:
                lip_file = os.path.join(output_folder, os.path.splitext(os.path.basename(audio_file))[0] + ".lip")
            voice_line = voice_lines[index]

            # Generate the lip file using FaceFXWrapper
            try:
               subprocess.run(["FaceFXWrapper.exe", game, language, fonix_data, audio_file, lip_file, voice_line])
            except FileNotFoundError:
               logging.exception("FaceFXWrapper.exe not found")
               messagebox.showerror("Error", "Missing file: FaceFXWrapper.exe not found.\nPlease place this next to the main .exe")
            except Exception as e:
               logging.exception("An error occurred during file processing")
               messagebox.showerror("Error", "An error occurred during file processing. Please check the log for details.")

               break # STOP SPAMMING ME WITH ERROR MESSAGES YOU CUNT

            completed_files += 1
            window.after(0, lambda: update_loading_bar(completed_files))

        # Enable the "Process Files" button once all the files are processed
        window.after(0, lambda: process_button.config(state=tk.NORMAL))
        messagebox.showinfo("Files Finished", "Task Complete! .lip files generated successfully.")

    # Start processing files in a separate thread
    threading.Thread(target=process_files_thread, daemon=True).start()

def import_dialog_file(file_path):
    pattern = r"([0-9A-F]{8}_\d)\.xwm"
    table_data = {}
    
    with open(file_path, 'r', encoding='utf-8-sig') as file:
        lines = file.readlines()
    
    for line in lines:
        match = re.match(r"([0-9A-F]{8}_\d)\.xwm\s+(.+)", line)
        if match:
            audio_file = match.group(1)
            voice_line = match.group(2).strip()
            table_data[audio_file] = {'Audio File': '', 'Voice Line': voice_line}
    
    # Update the table with voice lines
    for item in treeview.get_children():
        audio_file = treeview.item(item, "values")[0]
        file_name = os.path.splitext(os.path.basename(audio_file))[0]
        if file_name in table_data:
            table_data[file_name]['Audio File'] = audio_file
    
    # Update the treeview with the updated table data
    for item, data in table_data.items():
        treeview.insert("", "end", values=(data['Audio File'], data['Voice Line']))

def select_files(): # ---- Ironically THIS code gave me the most trouble ----
    audio_files = filedialog.askopenfilenames(filetypes=[("Audio Files", "*.wav")])
    audio_entry.delete(0, tk.END)
    audio_entry.insert(tk.END, ';'.join(audio_files))

    # Clear previous table data
    for item in treeview.get_children():
        treeview.delete(item)

    if dialog_export_file_path is None:
        continue_without_export = messagebox.askyesno(
            "Warning",
            "Without Exported Dialog from the Creation Kit, you may need to assign your lines manually!\n\nAre you sure you wish to continue?"
        )
        if not continue_without_export:
            return

    try:
        if dialog_export_file_path:
            process_dialog_export(dialog_export_file_path)

        # Populate the table with audio files and voice lines
        for audio_file in audio_files:
            audio_file_name = os.path.splitext(os.path.basename(audio_file))[0]
            voice_line = existing_table.get(audio_file_name, '')
            treeview.insert("", "end", values=(audio_file, voice_line))

    except FileNotFoundError: # Honestly how do you fuck up so bad to get this error anyways? I only added this for redundancy.
        messagebox.showerror("Error", "Exported Dialog file was not found.")
        logging.exception("Exported Dialogue not found")
    except Exception as e:
        logging.exception("An error occurred during file processing")
        messagebox.showerror("Error", "An error occurred during file processing. Please check the log for details.")

    dialog_export_file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
    if dialog_export_file_path:
        open_audio_files()

def open_audio_files():
    if dialog_export_file_path is None:
        continue_without_export = messagebox.askyesno(
            "Warning",
            "Without Exported Dialog from the Creation Kit, you may need to assign your lines manually!\n\nAre you sure you wish to continue?"
        )
        if not continue_without_export:
            return

    audio_files = filedialog.askopenfilenames(filetypes=[("Audio Files", "*.wav")])
    audio_entry.delete(0, tk.END)
    audio_entry.insert(tk.END, ';'.join(audio_files))

    # Clear previous table data
    for item in treeview.get_children():
        treeview.delete(item)

    try:
        if dialog_export_file_path:
            process_dialog_export(dialog_export_file_path)

        # Populate the table with audio files and voice lines
        for audio_file in audio_files:
            audio_file_name = os.path.splitext(os.path.basename(audio_file))[0]
            voice_line = existing_table.get(audio_file_name, '')
            treeview.insert("", "end", values=(audio_file, voice_line))

    except FileNotFoundError:
        messagebox.showerror("Error", "Exported Dialog file was not found.")

# def validate_file_layout(file_path): # --defunct code - we already built this code into process_dialog_export
#     try:
#         pattern = r"([0-9A-F]{8}_\d)\.xwm"
#         with open(file_path, 'r', encoding='utf-8-sig') as file:
#             reader = csv.DictReader(file, delimiter='\t')
#             for row in reader:
#                 full_path = row['FULL PATH']
#                 match = re.search(pattern, full_path)
#                 if not match:
#                     messagebox.showwarning(
#                         "Warning",
#                         "Invalid file layout found.\n\nPlease ensure the file follows the expected layout."
#                     )
#                     logging.warning(f"Invalid file layout found in file: {file_path}")
#                     return
#     except KeyError:
#         logging.warning(f"Invalid file layout found in file: {file_path}")
#         return # ---------------------------- Keeping this code handy for later

def process_dialog_export(file_path):
    try:
        pattern = r"([0-9A-F]{8}_\d)\.xwm"
        with open(file_path, 'r', encoding='utf-8-sig') as file:
            reader = csv.DictReader(file, delimiter='\t')
            for row in reader:
                full_path = row['FULL PATH']
                match = re.search(pattern, full_path)
                if match:
                    file_name = match.group(1)
                    voice_line = row['RESPONSE TEXT'].strip()
                    existing_table[file_name] = voice_line
                if not match:
                    messagebox.showwarning(
                        "Warning",
                        "Invalid file layout found.\n\nPlease ensure the file follows the expected layout."
                        )
                    logging.warning(f"Invalid file layout found in file: {file_path}")
                    print("Error reading file format - mismatched type")
                    return
    except KeyError: # What a pain this was to figure out
            logging.warning(f"Invalid file layout found in file: {file_path}")
            messagebox.showwarning(
                "Warning",
                "Invalid file layout found.\n\nPlease ensure the file follows the expected layout."
                )
            print("Error reading file format - Failed to parse dialogue.")
            return

# Assuming the existing table is a dictionary where the keys are file names and the values are voice lines
existing_table = {}

def update_voice_line(event): # Redundant code, only used here in case something breaks. will be removed in the next update
    try:
        selected_item = treeview.selection()[0]
        voice_line_entry.delete(0, tk.END)
        voice_line_entry.insert(tk.END, treeview.item(selected_item, "values")[1])
        voice_line_entry.focus()
        voice_line_entry.bind("<Return>", save_voice_line)
        voice_line_entry.bind("<FocusOut>", save_voice_line)
    except IndexError:
        print("HEY! You're supposed to add files before double clicking here, my guy!")

def save_voice_line(event): # Redundant code, only used here in case something breaks. will be removed in the next update
    try:
        global selected_item
        new_voice_line = voice_line_entry.get()
        treeview.set(selected_item, column="Voice Line", value=new_voice_line)
        voice_line_entry.unbind("<Return>")
        voice_line_entry.unbind("<FocusOut>")
        voice_line_entry.delete(0, tk.END)
    except NameError:
        print("Updating audio data table...\nIf this freezes or crashes, contact me via the forums or Discord and provide the error code.")

def open_dialog_export():
    global dialog_export_file_path
    if dialog_export_file_path:
        process_dialog_export(dialog_export_file_path)
        open_audio_files()
        return

    dialog_export_file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
    if dialog_export_file_path:
        process_dialog_export(dialog_export_file_path)
        open_audio_files()

def toggle_output_folder():
    if save_location_checkbox.get():
        output_entry.config(state="disabled")
    else:
        output_entry.config(state="normal")

def open_link():
    #webbrowser.open("https://ko-fi.com/eatingpizza")
    messagebox.showinfo("Authors Note", "Thank you for your support - every little bit counts.\n \nIf my tool helped you, consider leaving an endorsement on the Nexus page!")
	
def game_selection_changed(event):
    selected_game = game_var.get()
    if selected_game == 'Starfield':
        messagebox.showinfo("Starfield Release?", "WHEN IT'S READY.")
        game_var.set('Fallout4')

class EditableCell(tk.Entry): # I hope this fucking works I'm getting sick of this bullshit.
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.bind("<Return>", self.on_return)
        self.bind("<Escape>", self.on_escape)
        
    def on_return(self, event):
        self.master.focus()  # Remove focus from the Entry widget
        self.destroy()  # Destroy the Entry widget
        
    def on_escape(self, event):
        self.master.focus()  # Remove focus from the Entry widget
        self.delete(0, tk.END)  # Clear the Entry widget
        self.destroy()  # Destroy the Entry widget

def edit_cell(event): # THIS CODE IS A FUCKING NIGHTMARE I AM NEVER TOUCHING PYTHON AGAIN.
    try:
        item_id = treeview.identify_row(event.y)
        column = treeview.identify_column(event.x)
    
        if column != "#2":  # ONLY allow editing of the second column (Voice Line)
            return
    
        cell_value = treeview.item(item_id)["values"][1]
        bbox = treeview.bbox(item_id, column)
    
        # Create a separate window for editing
        edit_window = tk.Toplevel(window)
        edit_window.attributes("-topmost", True)
        edit_window.title("Edit Voice Text")
    
        # Calculate the position and size of the edit window
        x = window.winfo_x() + bbox[0]
        y = window.winfo_y() + bbox[1]
        width = max(bbox[2] - bbox[0], 300)  # Set minimum width to 300
        height = max(bbox[3] - bbox[1], 30)  # Set minimum height to 30
    
        edit_window.geometry(f"{width}x{height}+{x}+{y}")
    
        # Create an Entry widget in the edit window
        entry = ttk.Entry(edit_window, width=30)
        entry.pack(fill="both", expand=True)
        entry.insert(0, cell_value)
        entry.focus_set()
    
        def save_changes():
            new_value = entry.get()
            treeview.set(item_id, column=column, value=new_value)
            edit_window.destroy()
    
        def cancel_edit():
            edit_window.destroy()
    
        # Bind the Return key to save the changes
        entry.bind("<Return>", lambda event: save_changes())
        # Bind the Escape key to cancel the edit
        entry.bind("<Escape>", lambda event: cancel_edit())
    except IndexError:
        print("Stop clicking blank spaces! (Input Ignored)") # How the fuck do you even get this error?
        logging.exception("User clicked a blank space, but everything is okay.")
    except NameError:
        print("Updating audio data table...\nIf this freezes or crashes, contact me via the forums or Discord and provide the error code.")
        logging.exception("Updating audio data table...\nIf this freezes or crashes, contact me via the forums or Discord and provide the error code.")
# A NameError occurred during audio table cell editing. This will be reported to the gmod admins.

# Create the main window
window = tk.Tk()
window.title("Lip File Generator")
#style = ttkthemes.ThemedStyle(window)
#style.set_theme("smog")

# Create a frame for the setup settings
setup_frame = tk.Frame(window)
setup_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nw")
image_bytes = base64.b64decode(base64_image)
image = Image.open(io.BytesIO(image_bytes))

# Create a Tkinter PhotoImage from the image
photo = ImageTk.PhotoImage(image)

# Create a label to display the image
image_label = tk.Label(setup_frame, image=photo)
image_label.grid(row=3, column=0, columnspan=2)

# Create labels and entry fields for the setup settings
# Gonna use a dropdown box, might update this shit if that new game Skywind or whatever comes out. the space one
game_label = tk.Label(setup_frame, text="Game:")
game_label.grid(row=0, column=0, sticky="w")
game_var = tk.StringVar()
game_entry = ttk.Combobox(setup_frame, textvariable=game_var, state="readonly")
game_entry["values"] = ("Fallout4", "Skyrim", "Starfield")
game_entry.grid(row=0, column=1, pady=5)
game_entry.current(0)  # Set default value for Game
game_entry.bind("<<ComboboxSelected>>", game_selection_changed)

language_label = tk.Label(setup_frame, text="Language:")
language_label.grid(row=1, column=0, sticky="w")
language_entry = tk.Entry(setup_frame)
language_entry.grid(row=1, column=1, pady=5)
language_entry.insert(tk.END, "USEnglish")  # Set default value for Language

output_label = tk.Label(setup_frame, text="Output Folder:")
output_label.grid(row=2, column=0, sticky="w")
output_entry = tk.Entry(setup_frame)
output_entry.grid(row=2, column=1, pady=5)
output_description = tk.Label(window, text="Tool made by 54yeg. Special thanks to:\nThe XVASynth Dev Team, radbeetle, and ChatGPT")
output_description.grid(row=2, column=0)
link_label = tk.Label(window, text="If you found this helpful, consider buying me a coffee!", fg="blue", cursor="hand2")
link_label.bind("<Button-1>", lambda event: open_link())
link_label.grid(row=3, column=0)
# Okay here goes the checkbox
save_location_checkbox = tk.BooleanVar()
checkbox = tk.Checkbutton(setup_frame, text="Save in Same Location", variable=save_location_checkbox, command=toggle_output_folder)
checkbox.grid(row=3, column=0, columnspan=2, sticky="w", pady=(1, 250)) # DONT TOUCH THIS OR KILL

# Create a frame for the table view
table_frame = tk.Frame(window)
table_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

audio_label = tk.Label(table_frame, text="Audio Files:")
audio_label.pack()
audio_entry = tk.Entry(table_frame, state="readonly")
audio_entry.pack()

select_button = tk.Button(table_frame, text="Select Files", command=open_audio_files)
select_button.pack()
# Create the "Exported Dialog Processing" button
export_button = tk.Button(table_frame, text="Select Exported Dialogue", command=open_dialog_export)
export_button.pack()

export_description = tk.Label(table_frame, text="If you exported your Quests' Dialog from the Creation Kit, click this button to open\n your exported .txt file (usually in your Fallout 4 directory)")
export_description.pack()

# Create a treeview for displaying the table
treeview = ttk.Treeview(table_frame, columns=("Audio File", "Voice Line"), show="headings")
treeview.heading("Audio File", text="Audio File")
treeview.heading("Voice Line", text="Voice Line")
treeview.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)

# Create a scrollbar for the treeview
scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=treeview.yview)
treeview.configure(yscrollcommand=scrollbar.set)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Allow voice lines to be edited
treeview.bind("<Double-Button-1>", edit_cell) #update_voice_line, if edit_cell fucks up

# Create an Entry widget for editing voice lines
#voice_line_entry = tk.Entry(window) # this is for the deprecated update_voice_line code if I ever need to revert back
#voice_line_entry.grid(row=1, column=1, pady=5) # leaving it disabled for now

# Create a loading bar widget, and please fucking work
loading_bar = ttk.Progressbar(window, orient='horizontal', length=423, mode='determinate')
loading_bar.grid(row=1, column=1, padx=(0,17), pady=0) # MAKE ME COLUMN 1 AND ROW 1 OR I WILL EAT YOUR BRAINS

process_button = tk.Button(window, text="Process Files", command=process_files)
process_button.grid(row=2, column=1, pady=5)

process_description = tk.Label(window, text="Click this button to start the file processing.")
process_description.grid(row=3, column=1)

# Configure row and column weights to make them stretch with the window
window.grid_rowconfigure(0, weight=1)
window.grid_columnconfigure(1, weight=1)
table_frame.grid_rowconfigure(0, weight=1)
table_frame.grid_columnconfigure(0, weight=1)

# Create a frame for the buttons
button_frame = tk.Frame(table_frame)
button_frame.pack()

# Start the Tkinter event loop
window.mainloop()

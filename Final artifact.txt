import threading
import pyautogui
import tkinter as tk
from tkinter import messagebox, simpledialog, Tk
from tkinter.filedialog import askdirectory, askopenfilename
from docx import Document
from docx.shared import Inches, Cm
import os
import glob
from pynput import keyboard

# Global variable for screenshot count
screenshot_count = 0

# Function to be called when the Submit button is clicked
def on_ok():
    selected_option = var.get()
    if selected_option == 1:
        handle_option_1()
    elif selected_option == 2:
        handle_option_2()
    else:
        messagebox.showwarning("Selection Error", "Please select an option.")

# Function to be called when the Reset button is clicked
def on_reset():
    var.set(0)  # Reset the IntVar to its initial value (no selection)

def Close(): 
    confirm_exit = messagebox.askquestion("Exit", "Are you sure you want to exit?")
    if confirm_exit == "yes":
        root.destroy()

# Function to handle the second option (already created document)
def handle_option_2():
    messagebox.showinfo("Test Artifact Creation", "Test Artifact Creation with created document will start")

    # Prompt for already created Word document path
    doc_path = askopenfilename(title="Select the existing Word Document", filetypes=[("Word files", "*.docx")])
    if not doc_path:
        messagebox.showerror("Error", "No Word document selected.")
        return

    folder_path = askdirectory(title="Select folder to save screenshots").replace('/', '\\\\')
    if not folder_path:
        messagebox.showerror("Error", "No folder selected.")
        return

    # Run screenshot process in a separate thread to avoid freezing
    screenshot_thread = threading.Thread(target=run_screenshot_process, args=(doc_path, folder_path))
    screenshot_thread.start()

# Function to handle the first option (new document creation)
def handle_option_1():
    messagebox.showinfo("Test Artifact Creation", "Testing artifact creation with new Document will start")

    # Initializing directory path for folder, document name, and document file path
    folder_path = askdirectory(title="Select folder to save file").replace('/', '\\\\')
    doc_name = simpledialog.askstring(title="Document Name", prompt="Enter Document name for Artifact:\t\t\t\t\t\t")
    doc_path = folder_path + "\\\\" + doc_name + ".docx"

    # Run screenshot process in a separate thread to avoid freezing
    screenshot_thread = threading.Thread(target=run_screenshot_process_new, args=(doc_path, folder_path))
    screenshot_thread.start()

def run_screenshot_process(doc_path, folder_path):
    global screenshot_count
    screenshot_count = 0

    # Initialize the list to store screenshot paths
    screenshots = []

    # Initialize paths for saving screenshots
    ss_path_template = folder_path + "\\\\screenshot_{}.png"

    # Function to take a screenshot and add to memory (not saving immediately)
    def take_screenshot():
        global screenshot_count
        screenshot_count += 1
        screenshot = pyautogui.screenshot()
        screenshots.append(screenshot)
        print(f"Screenshot {screenshot_count} taken")

    # Function to batch write screenshots to disk and append to the Word document
    def append_screenshots_to_doc(screenshot_images, doc_path):
        # Optimize I/O: Batch save screenshots in one go after all screenshots are taken
        doc = Document(doc_path)
        sections = doc.sections
        for section in sections:
            section.top_margin = Cm(1)
            section.bottom_margin = Cm(1)
            section.left_margin = Cm(1)
            section.right_margin = Cm(1)

        # Save screenshots to disk in bulk
        for idx, screenshot in enumerate(screenshot_images):
            ss_path = ss_path_template.format(idx + 1)
            screenshot.save(ss_path)
            doc.add_picture(ss_path, width=Inches(7.5), height=Inches(4.5))

        # Save the document only once after all screenshots are appended
        doc.save(doc_path)
        print(f"Screenshots appended to the document at {doc_path}")

        # Cleanup: Remove screenshots from the folder
        png_files = glob.glob(os.path.join(folder_path, "*.png"))
        for file in png_files:
            try:
                os.remove(file)
            except Exception as e:
                print(f"Error deleting {file}: {e}")

    # Function to listen for the 'x' key for taking screenshots and 'esc' to stop
    def on_press(key):
        try:
            if key.char == 'x' or key.char == 'X':  # Check if 'x' key is pressed
                take_screenshot()
        except AttributeError:
            if key == keyboard.Key.esc:  # Stop the loop if 'esc' is pressed
                return False

    # Start screenshot process
    while True:
        result_1 = messagebox.askquestion("Take Screenshot", "Do you want to start taking screenshots?")
        if result_1 == "yes":
            print("Screenshot process started. Press 'X' to capture a screenshot. Press 'esc' to stop.")
            try:
                with keyboard.Listener(on_press=on_press) as listener:
                    listener.join()
            except KeyboardInterrupt:
                print("Loop interrupted via Esc.")

        # Asking whether the user still wants to continue taking screenshots
        result_2 = messagebox.askquestion("Pending Screenshots", "Do you want to continue taking screenshots?")
        if result_2 == "no":
            break

    # Batch process and save screenshots after exiting the loop
    if screenshots:
        append_screenshots_to_doc(screenshots, doc_path)
    else:
        print("No screenshots were taken.")
    print("Document is created. Process completed.")

def run_screenshot_process_new(doc_path, folder_path):
    global screenshot_count
    screenshot_count = 0

    # Initialize the list to store screenshot paths
    screenshots = []

    # Initialize paths for saving screenshots
    ss_path_template = folder_path + "\\\\screenshot_{}.png"

    # Function to take a screenshot and store in memory
    def take_screenshot():
        global screenshot_count
        screenshot_count += 1
        screenshot = pyautogui.screenshot()
        screenshots.append(screenshot)
        print(f"Screenshot {screenshot_count} taken")

    # Function to save screenshots to a new Word document (batch save)
    def save_screenshots_to_doc(screenshot_images, doc_path):
        doc = Document()
        sections = doc.sections
        for section in sections:
            section.top_margin = Cm(1)
            section.bottom_margin = Cm(1)
            section.left_margin = Cm(1)
            section.right_margin = Cm(1)

        # Save screenshots to disk and append them in one go
        for idx, screenshot in enumerate(screenshot_images):
            ss_path = ss_path_template.format(idx + 1)
            screenshot.save(ss_path)
            doc.add_picture(ss_path, width=Inches(7.5), height=Inches(4.5))

        # Save document only once after all screenshots are processed
        doc.save(doc_path)
        print(f"Document {doc_path} is ready with screenshots.")

        # Cleanup: Remove screenshots after appending
        png_files = glob.glob(os.path.join(folder_path, "*.png"))
        for file in png_files:
            try:
                os.remove(file)
            except Exception as e:
                print(f"Error deleting {file}: {e}")

    # Function to listen for the 'x' key for taking screenshots and 'esc' to stop
    def on_press(key):
        try:
            if key.char == 'x' or key.char == 'X':  # Check if 'x' key is pressed
                take_screenshot()
        except AttributeError:
            if key == keyboard.Key.esc:  # Stop the loop if 'esc' is pressed
                return False

    # Start screenshot process
    while True:
        result_1 = messagebox.askquestion("Take Screenshot", "Do you want to start taking screenshots?")
        if result_1 == "yes":
            print("Artifact creation will start. Press 'X' or 'x' to capture a screenshot. Press 'esc' to stop.")
            try:
                with keyboard.Listener(on_press=on_press) as listener:
                    listener.join()
            except KeyboardInterrupt:
                print("Loop interrupted via Esc.")

        # Asking whether the user still wants to continue taking screenshots
        result_2 = messagebox.askquestion("Pending Screenshots", "Do you want to continue taking screenshots?")
        if result_2 == "no":
            break

    # After all screenshots are taken, batch process and save them
    if screenshots:
        save_screenshots_to_doc(screenshots, doc_path)
    else:
        print("No screenshots were taken.")
    print("Document is created. Process completed.")

# Function to show the About section
def show_about():
    messagebox.showinfo("About", "This application is developed using Python.\n\nDeveloped for automating artifact creation with screenshots.")

# Main GUI setup
root = tk.Tk()
root.title("Artifact Automation")
root.geometry("500x450")

# Create a menu bar
menu_bar = tk.Menu(root)

# Add 'About' menu to the menu bar
about_menu = tk.Menu(menu_bar, tearoff=0)
about_menu.add_command(label="About", command=show_about)
menu_bar.add_cascade(label="Help", menu=about_menu)

# Set the menu bar to the root window
root.config(menu=menu_bar)

# Variable to store the selected radio button value
var = tk.IntVar()
var.set(0)  # Default value, none selected

# Creating radio buttons
radio1 = tk.Radiobutton(root, text="Test Artifact Creation with New Document", variable=var, value=1)
radio1.pack(anchor=tk.W, fill=tk.BOTH, padx=10, pady=20)

radio2 = tk.Radiobutton(root, text="Test Artifact Creation with already created document", variable=var, value=2)
radio2.pack(anchor=tk.W, fill=tk.BOTH, padx=10, pady=20)

# Frame to hold buttons in the same line
button_frame = tk.Frame(root)
button_frame.pack(pady=20)

# Submit button
sub_button = tk.Button(button_frame, text="Submit", command=on_ok)
sub_button.pack(side=tk.LEFT, padx=10)

# Reset button
reset_button = tk.Button(button_frame, text="Reset Selection", command=on_reset)
reset_button.pack(side=tk.LEFT, padx=10)

# Exit button
close_button = tk.Button(button_frame, text="Exit", command=Close)
close_button.pack(side=tk.LEFT, padx=10)

# Start the main loop
root.mainloop()

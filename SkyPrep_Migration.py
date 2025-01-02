import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl  # For reading and writing Excel files
import os
from datetime import datetime

def show_frame(frame):
    """Bring the selected frame to the front."""
    frame.tkraise()

def on_enter(e):
    """Change button color on hover to a darker version."""
    idx = button_widgets.index(e.widget)
    original_color = buttons[idx][1]
    # Darken the color slightly
    darker_color = "#%02x%02x%02x" % tuple(max(0, int(original_color[i:i+2], 16) - 30) for i in (1, 3, 5))
    e.widget['bg'] = darker_color

def on_leave(e):
    """Revert button color when hover ends."""
    idx = button_widgets.index(e.widget)
    e.widget['bg'] = buttons[idx][1]  # Original color

# Global variable to store the path of the uploaded file
uploaded_file_path = ""

def browse_file():
    """Browse and select an Excel file."""
    global uploaded_file_path
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    if file_path:
        uploaded_file_path = file_path
        file_label.config(text=f"Selected File: {os.path.basename(file_path)}")
    else:
        file_label.config(text="No file selected")

def start_cleanse():
    """Read the uploaded Excel file, process it, and save the result."""
    global uploaded_file_path
    if not uploaded_file_path:
        messagebox.showerror("Error", "Please upload an Excel file before starting.")
        return

    try:
        # Open the uploaded Excel file
        wb = openpyxl.load_workbook(uploaded_file_path)
        sheet = wb.active

        # Create a new workbook for the cleansed data
        new_wb = openpyxl.Workbook()
        new_sheet = new_wb.active

        # Get column indices based on headers
        headers = [cell.value for cell in sheet[1]]
        required_columns = [
            "Position ID",
            "Payroll Name",
            "Course Name Description",
            "Start Date",
            "Recertification Date",
            "Acquired Date",
        ]
        required_indices = [headers.index(col) for col in required_columns]

        # Write only the required headers to the new sheet
        new_sheet.append(required_columns)

        # Process each row after the header
        for row in sheet.iter_rows(min_row=2, values_only=True):
            row_list = list(row)  # Convert tuple to list to modify

            # Extract only the required columns for processing
            filtered_row = [row_list[idx] for idx in required_indices]

            # Apply cleansing rules to the filtered columns
            start_date = filtered_row[required_columns.index("Start Date")]
            recertification_date = filtered_row[required_columns.index("Recertification Date")]
            acquired_date = filtered_row[required_columns.index("Acquired Date")]

            # Rule 1: If there is a Start Date but no Recertification Date or Acquired Date, keep as is.
            if start_date and not recertification_date and not acquired_date:
                pass  # No change needed

            # Rule 2: If Start Date and Acquired Date exist, but no Recertification Date, clear both fields.
            elif start_date and acquired_date and not recertification_date:
                filtered_row[required_columns.index("Recertification Date")] = None
                filtered_row[required_columns.index("Acquired Date")] = None

            # Rule 3: If Start Date and Recertification Date exist, check their values.
            elif start_date and recertification_date:
                if recertification_date > start_date:
                    # Write Start Date as Acquired Date
                    filtered_row[required_columns.index("Acquired Date")] = start_date
                elif recertification_date == start_date:
                    # Do not write anything in Recertification Date and Acquired Date
                    filtered_row[required_columns.index("Recertification Date")] = None
                    filtered_row[required_columns.index("Acquired Date")] = None
                elif recertification_date < start_date:
                    # Clear both Recertification Date and Acquired Date
                    filtered_row[required_columns.index("Recertification Date")] = None
                    filtered_row[required_columns.index("Acquired Date")] = None

            # Append the processed filtered row to the new sheet
            new_sheet.append(filtered_row)

        # Ask the user where to save the new workbook
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Save Cleansed Data",
        )
        if not save_path:
            messagebox.showinfo("Cancelled", "Save operation was cancelled.")
            return

        # Save the new workbook to the chosen path
        new_wb.save(save_path)
        messagebox.showinfo("Success", f"Cleansed data saved to: {save_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Main window setup
root = tk.Tk()
root.title("SkyPrep Migration Tool")
root.geometry("600x400")
root.minsize(600, 400)  # Set minimum size to prevent distortion
root.configure(bg="#2E2E2E")  # Background color

# Main container frames
main_frame = tk.Frame(root)
main_frame.pack(fill="both", expand=True)

# Left menu frame
menu_frame = tk.Frame(main_frame, width=150, bg="#3C3F41", relief="raised")
menu_frame.pack(side="left", fill="y")

# Right content frame
content_frame = tk.Frame(main_frame, bg="#F5F5F5")
content_frame.pack(side="right", expand=True, fill="both")

# Bottom bar
bottom_bar = tk.Frame(root, bg="#2E2E2E", height=50)
bottom_bar.pack(side="bottom", fill="x")

# Footer label with dynamic text
footer_label = tk.Label(
    bottom_bar,
    text=f"© Voyago | {datetime.now().strftime('%Y-%m-%d')}",
    bg="#2E2E2E",
    fg="white",
    font=("Arial", 10),
)
footer_label.pack(side="right", padx=10)

# Define frames for each screen in the content area
cleanse_frame = tk.Frame(content_frame, bg="#F5F5F5")
transform_frame = tk.Frame(content_frame, bg="#F5F5F5")
compare_frame = tk.Frame(content_frame, bg="#F5F5F5")
transfer_frame = tk.Frame(content_frame, bg="#F5F5F5")

# Place all frames on the same stack
for frame in (cleanse_frame, transform_frame, compare_frame, transfer_frame):
    frame.place(relwidth=1, relheight=1)

# Add widgets to the Cleanse Screen
tk.Label(cleanse_frame, text="Cleanse Section", bg="#F5F5F5", font=("Arial", 16)).pack(pady=10)

file_button = tk.Button(cleanse_frame, text="Browse", font=("Arial", 12), command=browse_file)
file_button.pack(pady=5)

file_label = tk.Label(cleanse_frame, text="No file selected", bg="#F5F5F5", font=("Arial", 10))
file_label.pack(pady=5)

start_button = tk.Button(cleanse_frame, text="Start Cleanse", font=("Arial", 12), command=start_cleanse)
start_button.pack(pady=10)

# Add widgets to other screens (optional content placeholders)
tk.Label(transform_frame, text="Transform Screen", bg="#F5F5F5", font=("Arial", 16)).pack(pady=20)
tk.Label(compare_frame, text="Compare Screen", bg="#F5F5F5", font=("Arial", 16)).pack(pady=20)
tk.Label(transfer_frame, text="Transfer Screen", bg="#F5F5F5", font=("Arial", 16)).pack(pady=20)

# Function to dynamically resize buttons with spacing and padding
def resize_buttons():
    """Set rectangular buttons with uniform size."""
    frame_width = menu_frame.winfo_width()
    button_width = frame_width - 2 * padding  # Button width matches the menu frame width with padding
    button_height = 60  # Fixed height for all buttons

    for idx, button in enumerate(button_widgets):
        button.place(
            x=padding,  # Align with padding
            y=padding + idx * (button_height + spacing),  # Space buttons vertically
            width=button_width,
            height=button_height,
        )

# Define button properties
buttons = [(text, "#E90000", frame) for text, frame in [
    ("Cleanse", cleanse_frame),
    ("Transform", transform_frame),
    ("Compare", compare_frame),
    ("Transfer", transfer_frame),
]]
button_widgets = []
padding = 20  # Padding around the buttons
spacing = 30  # Space between buttons

# Add buttons to the menu frame with hover effects
for text, color, frame in buttons:
    btn = tk.Button(
        menu_frame,
        text=text,
        bg=color,
        fg="white",
        font=("Arial", 12, "bold"),
        relief="raised",
        borderwidth=2,
        command=lambda f=frame: show_frame(f),
    )
    btn.bind("<Enter>", on_enter)
    btn.bind("<Leave>", on_leave)
    button_widgets.append(btn)

# Bind the resize event to dynamically adjust button size, padding, and spacing
menu_frame.bind("<Configure>", lambda e: resize_buttons())

# Show the first screen by default
show_frame(cleanse_frame)

# Run the application
root.mainloop()

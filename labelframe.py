import tkinter as tk
from tkinter import ttk

# Create a function to be called when the button is clicked
def button_click():
    print("Button clicked!")

# Create the root window
root = tk.Tk()
root.title("LabelFrame Example")

# Create a LabelFrame
label_frame = ttk.Labelframe(root, text="LabelFrame")

# Add a button inside the LabelFrame
button = ttk.Button(label_frame, text="Click Me", command=button_click)
button.grid(row=0, column=0, padx=10, pady=10)

# Add more widgets if needed

# Place the LabelFrame within the root window
label_frame.grid(row=0, column=0, padx=20, pady=60)

# Start the Tkinter event loop
root.mainloop()

import tkinter as tk
from tkinter import ttk

root = tk.Tk()
# root.geometry("400x300")

# Create a Text widget
text_widget = tk.Text(root, wrap="word")
text_widget.pack(side="left", fill="both", expand=True)

# Create a Scrollbar widget and attach it to the Text widget
scrollbar = ttk.Scrollbar(root, orient="vertical", command=text_widget.yview)
scrollbar.pack(side="right", fill="y")
text_widget.config(yscrollcommand=scrollbar.set)

# Add some text to the Text widget
for i in range(1, 100):
    text_widget.insert("end", f"Line {i}\n")

root.mainloop()

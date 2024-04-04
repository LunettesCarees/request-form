import tkinter as tk

def checkbox_clicked(event, checkbox_index):
    # Introduce a slight delay before printing the value
    root.after(1000, lambda: print(f"Checkbox {checkbox_index+1} value:", checkbox_vars[checkbox_index].get()))

root = tk.Tk()

# Create IntVar variables for each checkbox
checkbox_vars = [tk.IntVar() for _ in range(4)]

# Labels for the checkboxes
checkbox_labels = ["Checkbox 1", "Checkbox 2", "Checkbox 3", "Checkbox 4"]

# Create checkboxes
checkboxes = []
for i in range(4):
    checkbox = tk.Checkbutton(root, text=checkbox_labels[i], variable=checkbox_vars[i])
    checkbox.bind("<Button-1>", lambda event, index=i: checkbox_clicked(event, index))
    checkbox.pack()
    checkboxes.append(checkbox)

root.mainloop()

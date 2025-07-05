import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from xml2rtf.converter import convert

def load_file():
    file_path = filedialog.askopenfilename(filetypes=[("XML Files", "*.xml")])
    if file_path:
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                xml_text = file.read()
                rtf_output = convert(xml_text)
                output_box.delete("1.0", tk.END)
                output_box.insert(tk.END, rtf_output)
                save_button.config(state=tk.NORMAL)
                root.rtf_data = rtf_output
        except Exception as e:
            messagebox.showerror("Error", str(e))

def save_file():
    file_path = filedialog.asksaveasfilename(defaultextension=".rtf", filetypes=[("RTF Files", "*.rtf")])
    if file_path:
        try:
            with open(file_path, "w", encoding="utf-8") as file:
                file.write(root.rtf_data)
            messagebox.showinfo("Success", "RTF file saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", str(e))

root = tk.Tk()
root.title("XML to RTF Converter")
root.geometry("800x600")
root.rtf_data = ""

frame = tk.Frame(root)
frame.pack(padx=10, pady=10, fill="both", expand=True)

tk.Label(frame, text="Converted RTF Output:").pack(anchor="w")

output_box = scrolledtext.ScrolledText(frame, wrap=tk.WORD, font=("Courier", 10))
output_box.pack(fill="both", expand=True)

button_frame = tk.Frame(root)
button_frame.pack(pady=10)

tk.Button(button_frame, text="Load XML", command=load_file).pack(side="left", padx=5)
save_button = tk.Button(button_frame, text="Save RTF", command=save_file, state=tk.DISABLED)
save_button.pack(side="left", padx=5)

root.mainloop()

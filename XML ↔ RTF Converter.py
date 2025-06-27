import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from lxml import etree

# XML to RTF Conversion Logic 
def xml_to_rtf(xml_text):
    try:
        root = etree.fromstring(xml_text.encode('utf-8'))

        def rtf_format(text, attrib):
            if not text:
                return ""
            rtf = ""
            if attrib.get("bold") == "true":
                rtf += r"\b "
            if attrib.get("italic") == "true":
                rtf += r"\i "
            if attrib.get("sup") == "true":
                rtf += r"\super "
            if attrib.get("sub") == "true":
                rtf += r"\sub "

            rtf += text

            if attrib.get("bold") == "true":
                rtf += r" \b0"
            if attrib.get("italic") == "true":
                rtf += r" \i0"
            if attrib.get("sup") == "true" or attrib.get("sub") == "true":
                rtf += r" \nosupersub"
            return rtf

        def parse_paragraph(p):
            rtf_line = ""
            is_heading = p.attrib.get("style") == "Title_document"

            if is_heading:
                rtf_line += r"\fs36\b "

            for r in p.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                    rtf_line += rtf_format(t.text.strip(), r.attrib)

            if is_heading:
                rtf_line += r" \b0\fs24"

            return rtf_line + r"\par\n"

        rtf_output = ""
        if root.tag.lower() == "p":
            rtf_output += parse_paragraph(root)
        else:
            for p in root.findall(".//p"):
                rtf_output += parse_paragraph(p)

        return "{\\rtf1\\ansi\n" + rtf_output + "}"
    except Exception as e:
        return f"Error parsing XML: {e}"

# GUI Functionalities 
def open_xml_file():
    file_path = filedialog.askopenfilename(filetypes=[("XML Files", "*.xml")])
    if not file_path:
        return
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            xml_content = file.read()
            xml_input.delete("1.0", tk.END)
            xml_input.insert(tk.END, xml_content)
    except Exception as e:
        messagebox.showerror("Error", str(e))

def convert_and_preview():
    xml_text = xml_input.get("1.0", tk.END).strip()
    if not xml_text:
        messagebox.showwarning("Empty Input", "XML content is empty.")
        return
    try:
        rtf_result = xml_to_rtf(xml_text)
        rtf_output.delete("1.0", tk.END)
        rtf_output.insert(tk.END, rtf_result)
    except Exception as e:
        messagebox.showerror("Conversion Error", str(e))

def save_rtf():
    rtf_text = rtf_output.get("1.0", tk.END).strip()
    if not rtf_text:
        messagebox.showwarning("Empty Output", "RTF content is empty.")
        return
    file_path = filedialog.asksaveasfilename(defaultextension=".rtf", filetypes=[("RTF Files", "*.rtf")])
    if not file_path:
        return
    try:
        with open(file_path, "w", encoding="utf-8") as file:
            file.write(rtf_text)
        messagebox.showinfo("Saved", f"RTF saved to {file_path}")
    except Exception as e:
        messagebox.showerror("Save Error", str(e))



# GUI Layout 
root = tk.Tk()
root.title("XML to RTF Converter")
root.geometry("900x700")
root.resizable(True, True)

# Top Buttons
top_frame = tk.Frame(root)
top_frame.pack(pady=10)

btn_select = tk.Button(top_frame, text="Open XML File", command=open_xml_file, bg="lightblue", fg="black")
btn_select.grid(row=0, column=0, padx=10)

btn_convert = tk.Button(top_frame, text="Convert to RTF", command=convert_and_preview, bg="lightgreen", fg="black")
btn_convert.grid(row=0, column=1, padx=10)

btn_save = tk.Button(top_frame, text="Save RTF", command=save_rtf, bg="lightcoral", fg="black")
btn_save.grid(row=0, column=2, padx=10)

# Input XML
tk.Label(root, text="XML Input:").pack()
xml_input = scrolledtext.ScrolledText(root, height=12)
xml_input.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

# Output RTF
tk.Label(root, text="RTF Output:").pack()
rtf_output = scrolledtext.ScrolledText(root, height=12)
rtf_output.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

root.mainloop()

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from lxml import etree

# XML to RTF Conversion Logic 
def xml_to_rtf(xml_text):
    try:
        root = etree.fromstring(xml_text.encode('utf-8'))

#parse funtions

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
            if style in ["journaltitle", "articletitle"]:
                rtf += r" \i0"
            if style == "credittaxonomy":
                rtf += r" \b0"    
            return rtf



        def parse_paragraph(p):
            rtf = ""
            for r in p.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                    rtf += rtf_format(t.text.strip(), r.attrib)
            rtf += r"\par\n"
            return rtf
        
        def parse_supp_media(p):
            rtf = r"\fs22 Supplementary: "
            for r in p.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                    rtf += rtf_format(t.text.strip(), r.attrib)
            rtf += r" \fs24\par\n"
            return rtf

        def parse_referenced_data(p):
            rtf = r"\fs22\ul "
            for r in p.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                    rtf += rtf_format(t.text.strip(), r.attrib)
            rtf += r" \ul0\fs24\par\n"
            return rtf

        def parse_bib_entry(p):
            rtf = r"{\pard\box\brdrdot "
            for r in p.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                    rtf += rtf_format(t.text.strip(), r.attrib)
            rtf += r" \par}"
            return rtf + "\n"


        def parse_title(p):
            rtf = r"\fs48\b\qc"
            for r in p.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                    rtf += rtf_format(t.text.strip(), r.attrib)
            rtf += r"\b0\fs48\qc0\par\n"   
            return rtf
                 
                 
        def parse_authors(p):
            rtf = r"\fs24\i "
            for r in p.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                    rtf += rtf_format(t.text.strip(), r.attrib)
            rtf += r"\i0\par\n"
            return rtf
        
        
        def parse_abstract(p):
            rtf = r"\fs22 "
            for r in p.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                    rtf += rtf_format(t.text.strip(), r.attrib)
            rtf += r" \fs24\par\n"
            return rtf
        

        def parse_list(p):
            rtf = r"\bullet\tab "
            for r in p.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                    rtf += rtf_format(t.text.strip(), r.attrib)
            rtf += r"\par\n"
            return rtf     
        
        
        def parse_heading(p):
            rtf = r"\fs36\b"
            for r in p.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                    rtf += rtf_format(t.text.strip(), r.attrib)
            rtf += r"\b0\fs24\par\n"
            return rtf  


        def parse_keywords(p):
            rtf = r"\i Keywords: "
            for r in p.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                    rtf += t.text.strip() + " "
            rtf += r"\i0\par\n"
            return rtf
        
        
        def parse_affiliation(p):
            rtf = r"\fs22\tab "
            for r in p.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                    rtf += t.text.strip() + " "
            rtf +=r"\fs24\par\n"
            return rtf         
        
        def parse_abshead(p):
            rtf = r"\fs28\b"
            for r in p.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                    rtf += t.text.strip() + " "
            rtf +=r"\b0\fs24\par\n"
            return rtf     
        
        def parse_email(p):
            rtf = r"\cf1\ul "
            for r in p.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                    rtf += t.text.strip() + " "
            rtf +=r"\ul0\cf0\par\n"
            return rtf     
        
        def parse_hyperlink(hyperlink):
            Url = hyperlink.attrib.get("Url", "")
            display_text = ""
            
            for r in hyperlink.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                    display_text += rtf_format(t.text.strip(), r.attrib)
            return (
                r"{\field{\*\fldinst{HYPERLINK \"" + Url + r"\"}}"
                r"{\fldrslt{\cf1\ul " + display_text + r"\ul0\cf0}}}"
            )          

        rtf_output = ""
        if root.tag.lower() == "p":
            rtf_output += parse_paragraph(root)
        else:
            for p in root.findall(".//p"):
                style = p.attrib.get("style", "").lower()
                
                if style == "title_document":
                    rtf_output += parse_title(p)
                elif style == "authors":
                    rtf_output += parse_authors(p)
                elif style == "abstract":
                    rtf_output += parse_abstract(p)
                elif style == "list paragraph":
                    rtf_output += parse_list(p)
                elif style in ["head1", "heading1"]:
                    rtf_output += parse_heading(p)
                elif style =="keywords" :
                    rtf_output += parse_keywords(p)  
                elif style =="affiliation":
                    rtf_output += parse_affiliation(p)    
                elif style == "abshead":
                    rtf_output += parse_abshead(p)    
                elif style == "email":
                    rtf_output += parse_email(p)     
                elif style == "suppmedia":
                    rtf_output += parse_supp_media(p)
                elif style == "referenceddata":
                    rtf_output += parse_referenced_data(p)
                elif style == "bib_entry":
                    rtf_output += parse_bib_entry(p)    
                    
                    
                else: 
                    rtf_output += parse_paragraph(p) 
                    
                               
            for el in root.iter("Hyperlink"):
                rtf_output += parse_hyperlink(el) + r"\par\n"   
                
        

        return "{\\rtf1\\ansi\\deff0{\\colortbl;\\red0\\green0\\blue255;}\n" + rtf_output + "}"
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

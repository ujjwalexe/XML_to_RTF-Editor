import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from lxml import etree
import subprocess   
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

        bib_counter = [0] #Counter 
        
        def parse_bib_entry(p):
            bib_counter[0] += 1
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
            rtf = r"\li200\bullet\tab "
            for r in p.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                    rtf += rtf_format(t.text.strip(), r.attrib)
            rtf += r"\li0\par\n"
            return rtf     
        
        heading_counter = [0] #Counter
        
        def parse_heading(p):
            heading_counter[0] += 1
            rtf = r"\fs36\b {heading_counter[0]}"
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
        
        def parse_article_title(p):
            rtf = r"\fs36\b\qc "
            for r in p.findall("r"):
                t = r.find("t")
            if t is not None and t.text:
                rtf += rtf_format(t.text.strip(), r.attrib)
            rtf += r"\b0\qc0\fs24\par\n"
            return rtf

        def parse_journaltitle(p):
            rtf = r"\fs24\i\qc "
            for r in p.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                     rtf += rtf_format(t.text.strip(), r.attrib)
            rtf += r"\i0\qc0\par\n"
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
        def parse_booktitle(p):
            rtf = r"\fs24\b\i "
            for r in p.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                    rtf += rtf_format(t.text.strip(), r.attrib)
            rtf += r" \i0\b0\par\n"
            return rtf
        
        
        def parse_doi_url(p):
            rtf = r"\cf1\ul "
            for r in p.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                 rtf += t.text.strip() + " "
            rtf += r"\ul0\cf0\par\n"
            return rtf

        def parse_pubinfo(p):
            rtf = r"\fs22\tab "
            for r in p.findall("r"):
                t = r.find("t")
            if t is not None and t.text:
                rtf += t.text.strip() + " "
            rtf += r"\fs24\par\n"
            return rtf
       
        def parse_subheading(p):
            rtf = r"\fs28\b "
            for r in p.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                    rtf += t.text.strip() + " "
            rtf += r"\b0\fs24\par\n"
            return rtf
        
        def parse_affiliation_detail(p):
            rtf = r"\fs22\tab "
            for r in p.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                    rtf += t.text.strip() + " "
            rtf += r"\fs24\par\n"
            return rtf
    
        
        figure_counter = [0] #counter
        table_counter = [0] #counter

        def parse_caption(p, caption_type):
            if caption_type == "figure":
                figure_counter[0] += 1
                prefix = f"Figure {figure_counter[0]}"
            else:
                table_counter[0] += 1
                prefix = f"Table {table_counter[0]}"

            rtf = rf"\fs22\i {prefix}: "
            for r in p.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                    rtf += rtf_format(t.text.strip(), r.attrib)
            rtf += r" \i0\fs24\par\n"
            return rtf


        def parse_label(p):
            rtf = r"\b "
            for r in p.findall("r"):
                t = r.find("t")
                if t is not None and t.text:
                    rtf += t.text.strip() + " "
            rtf += r"\b0\par\n"
            return rtf


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
                elif style == "articletitle":
                    rtf_output += parse_article_title(p)
                elif style == "journaltitle":
                    rtf_output += parse_journaltitle(p)   
                elif style == "booktitle":
                    rtf_output += parse_booktitle(p)
                elif style in ["doi", "url"]:
                    rtf_output += parse_doi_url(p)
                elif style in ["year", "volume", "issue", "pages"]:
                    rtf_output += parse_pubinfo(p)
                elif style in ["head2", "ackhead", "conflictofinteresthead", "referencehead", "keywordhead"]:
                    rtf_output += parse_subheading(p)
                elif style in ["orgname", "orgdiv", "city", "state", "country", "pincode", "street"]:
                    rtf_output += parse_affiliation_detail(p)
                elif style == "figurecaption":
                    rtf_output += parse_caption(p, "figure")
                elif style == "tablecaption":
                    rtf_output += parse_caption(p, "table")

                elif style == "label":
                    rtf_output += parse_label(p)        
                elif style in ["ackpara", "conflictofinterest", "refmisc"]:
                    rtf_output += parse_paragraph(p)

                    
                else: 
                    rtf_output += parse_paragraph(p) 
                    
                               
            for el in root.iter("Hyperlink"):
                rtf_output += parse_hyperlink(el) + r"\par\n"   
                
        

        return (
                    r"{\rtf1\ansi\deff0"
                    r"{\colortbl;\red0\green0\blue255;}"
                    "\n" + rtf_output + "}"
        )
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
            xml_input_tab.delete("1.0", tk.END)
            xml_input_tab.insert(tk.END, xml_content)
    except Exception as e:
        messagebox.showerror("Error", str(e))

def convert_and_preview():
    xml_text = xml_input_tab.get("1.0", tk.END).strip()
    if not xml_text:
        messagebox.showwarning("Empty Input", "XML content is empty.")
        return
    try:
        rtf_result = xml_to_rtf(xml_text)
        rtf_output_tab.delete("1.0", tk.END)
        rtf_output_tab.insert(tk.END, rtf_result)
    except Exception as e:
        messagebox.showerror("Conversion Error", str(e))

def save_rtf():
    rtf_text = rtf_output_tab.get("1.0", tk.END).strip()
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

def save_and_open_rtf():
    rtf_text = rtf_output_tab.get("1.0", tk.END).strip()
    if not rtf_text:
        messagebox.showwarning("Empty Output", "RTF content is empty.")
        return
    file_path = filedialog.asksaveasfilename(defaultextension=".rtf", filetypes=[("RTF Files", "*.rtf")])
    if not file_path:
        return
    try:
        with open(file_path, "w", encoding="utf-8") as file:
            file.write(rtf_text)
        subprocess.Popen([r"C:\Users\A9790\Desktop\TextEditor-master\TextEditor-master\Text Editor\bin\Debug\Text Editor.exe", file_path], shell=True)
    except Exception as e:
        messagebox.showerror("Error Opening File", str(e))

# GUI Layout
root = tk.Tk()
root.title("XML to RTF Converter")
root.geometry("1000x700")
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

btn_save_open = tk.Button(top_frame, text="Save & Open RTF", command=save_and_open_rtf, bg="orange", fg="black")
btn_save_open.grid(row=0, column=3, padx=10)

# Tabbed Layout
tab_control = ttk.Notebook(root)

xml_input_tab = scrolledtext.ScrolledText(tab_control, height=30)
tab_control.add(xml_input_tab, text="XML Input")

rtf_output_tab = scrolledtext.ScrolledText(tab_control, height=30)
tab_control.add(rtf_output_tab, text="RTF Output")

tab_control.pack(expand=1, fill='both', padx=10, pady=10)

root.mainloop()

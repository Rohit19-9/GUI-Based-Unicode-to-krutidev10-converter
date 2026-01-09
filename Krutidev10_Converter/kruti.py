import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

# ==========================================================
# OFFICIAL UNICODE → KRUTIDEV10 ENGINE (NO CHANGES)
# ==========================================================

def Unicode_to_KrutiDev(unicode_substring):

    modified_substring = unicode_substring

    array_one = [
        "‘","’","“","”","(",")","{","}","=","।","?","-","µ","॰",",",".","् ",
        "०","१","२","३","४","५","६","७","८","९","x",

        "फ़्","क़","ख़","ग़","ज़्","ज़","ड़","ढ़","फ़","य़","ऱ","ऩ",
        "त्त्","त्त","क्त","दृ","कृ",

        "ह्न","ह्य","हृ","ह्म","ह्र","ह्","द्द","क्ष्","क्ष","त्र्","त्र","ज्ञ",
        "छ्य","ट्य","ठ्य","ड्य","ढ्य","द्य","द्व",
        "श्र","ट्र","ड्र","ढ्र","छ्र","क्र","फ्र","द्र","प्र","ग्र","रु","रू",
        "्र",

        "ओ","औ","आ","अ","ई","इ","उ","ऊ","ऐ","ए","ऋ",

        "क्","क","क्क","ख्","ख","ग्","ग","घ्","घ","ङ",
        "चै","च्","च","छ","ज्","ज","झ्","झ","ञ",

        "ट्ट","ट्ठ","ट","ठ","ड्ड","ड्ढ","ड","ढ","ण्","ण",
        "त्","त","थ्","थ","द्ध","द","ध्","ध","न्","न",

        "प्","प","फ्","फ","ब्","ब","भ्","भ","म्","म",
        "य्","य","र","ल्","ल","ळ","व्","व",
        "श्","श","ष्","ष","स्","स","ह",

        "ऑ","ॉ","ो","ौ","ा","ी","ु","ू","ृ","े","ै",
        "ं","ँ","ः","ॅ","ऽ","् ","्"
    ]

    array_two = [
        "^","*","Þ","ß","¼","½","¿","À","¾","A","\\","&","&","Œ","]","-","~ ",
        "å","ƒ","„","…","†","‡","ˆ","‰","Š","‹","Û",

        "¶","d","[k","x","T","t","M+","<+","Q",";","j","u",
        "Ù","Ùk","Dr","–","—",

        "à","á","â","ã","ºz","º","í","{","{k","«","=","K",
        "Nî","Vî","Bî","Mî","<î","|","}",
        "J","Vª","Mª","<ªª","Nª","Ø","Ý","æ","ç","xz","#",":",
        "z",

        "vks","vkS","vk","v","bZ","b","m","Å",",s",",","_",

        "D","d","ô","[","[k","X","x","?","?k","³",
        "pkS","P","p","N","T","t","÷",">","¥",

        "ê","ë","V","B","ì","ï","M","<",".",".k",
        "R","r","F","Fk",")","n","/","/k","U","u",

        "I","i","¶","Q","C","c","H","Hk","E","e",
        "¸",";","j","Y","y","G","O","o",
        "'","'k","\"","\"k","L","l","g",

        "v‚","‚","ks","kS","k","h","q","w","`","s","S",
        "a","¡","%","W","·","~ ","~"
    ]

    # preprocessing
    modified_substring = modified_substring.replace("ि","f")

    for i in range(len(array_one)):
        modified_substring = modified_substring.replace(array_one[i], array_two[i])

    # reposition f
    modified_substring = "  " + modified_substring + "  "
    pos = modified_substring.find("f")
    while pos != -1:
        modified_substring = (
            modified_substring[:pos-1]
            + "f"
            + modified_substring[pos-1]
            + modified_substring[pos+1:]
        )
        pos = modified_substring.find("f", pos+1)
    return modified_substring.strip()

# ==========================================================
# GUI + EXCEL
# ==========================================================

def browse_file():
    file = filedialog.askopenfilename(filetypes=[("Excel Files","*.xlsx *.xls")])
    if file:
        entry.delete(0, tk.END)
        entry.insert(0, file)
        df = pd.read_excel(file, dtype=str)
        menu["menu"].delete(0,"end")
        for col in df.columns:
            menu["menu"].add_command(label=col, command=tk._setit(selected_col,col))
        selected_col.set(df.columns[0])

def convert_excel():
    try:
        file = entry.get()
        if not file:
            messagebox.showerror("Error","Select Excel file")
            return

        df = pd.read_excel(file, dtype=str)
        col = selected_col.get()

        df[col+"_KrutiDev10"] = df[col].apply(
            lambda x: Unicode_to_KrutiDev(x) if isinstance(x,str) else x
        )

        out = os.path.splitext(file)[0] + "_krutidev.xlsx"
        df.to_excel(out,index=False)
        messagebox.showinfo("Success",f"Converted Successfully\n\n{out}")

    except Exception as e:
        messagebox.showerror("Error",str(e))

# ==========================================================
# GUI LAYOUT
# ==========================================================

root = tk.Tk()
root.title("Unicode → KrutiDev10 (Government Engine)")
root.geometry("560x280")
root.resizable(False,False)

tk.Label(root,text="Unicode to KrutiDev10 Converter",
         font=("Segoe UI",14,"bold")).pack(pady=10)

frame = tk.Frame(root)
frame.pack(padx=20,pady=20,fill="x")

tk.Label(frame,text="Excel File").grid(row=0,column=0,sticky="w")
entry = tk.Entry(frame,width=45)
entry.grid(row=0,column=1,padx=5)
tk.Button(frame,text="Browse",command=browse_file).grid(row=0,column=2)

tk.Label(frame,text="Column").grid(row=1,column=0,sticky="w",pady=10)
selected_col = tk.StringVar()
menu = tk.OptionMenu(frame,selected_col,"")
menu.config(width=32)
menu.grid(row=1,column=1,sticky="w")

tk.Button(frame,text="Convert",
          bg="green",fg="white",
          command=convert_excel,width=25).grid(row=2,column=1,pady=20)

root.mainloop()

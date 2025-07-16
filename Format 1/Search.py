import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import time as t

t.sleep(0.25)
print("""_______________________________________________ All Logs _______________________________________________""")

t.sleep(0.25)
print(">>> Searching Excel File...")

file_path = 'Test_File.xlsx' # File path to the Excel file

t.sleep(0.25)
print(">>> Excel File Found! " + " | " + file_path + " | " + " Is Name Of File And Ready To Be Opened!")

t.sleep(0.5)
print(">>> Opening Excel File In BG...")
t.sleep(0.5)
print(">>> Opening Search Feature Made By VPX...")
t.sleep(0.5)
print(">>> Last Step Assembling GUI And Excel File... Thank you For Your Patience!")
t.sleep(0.5)
print(">>> Opened GUI And Assembled All Data Successfully! Enjoy Searching! :)")

df = pd.read_excel(file_path, header=2)
df.columns = df.columns.str.strip() 
df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
df = df.fillna('-')

for col in ['Academy Roll No.', 'Sr. No.']:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

column_map = {
    'Name': 'Name',
    'Father Name': 'Father Name',
    'Roll No': 'Academy Roll No.',
    'Date of Birth': 'Date Of Birth',
    'Mobile No': 'Mob. No.',
    'Admit Card No': 'Admit Card No.',
    'Written Status': 'Written Status',
    'Physical Status': 'Physical Status',
    'Medical Status': 'Medical Status',
    'Remarks': 'Remarks'
}

root = tk.Tk()
root.title("Excel Se Data Nikaalne Ka Software")
root.geometry("500x500")

tk.Label(root, text="Select Field to Search:", font=("Arial", 12)).pack(pady=10)
selected_field = tk.StringVar()
selected_field.set("Name")
dropdown = ttk.Combobox(root, textvariable=selected_field, values=list(column_map.keys()), state="readonly", font=("Arial", 11))
dropdown.pack(pady=5)

tk.Label(root, text="Enter text to search:", font=("Arial", 12)).pack(pady=10)
search_entry = tk.Entry(root, width=50, font=("Arial", 12))
search_entry.pack(pady=5)

result_frame = ttk.Frame(root)
result_frame.pack(fill="both", expand=True, padx=10, pady=20)

tree_scroll_y = ttk.Scrollbar(result_frame, orient='vertical')
tree_scroll_y.pack(side="right", fill="y")

tree_scroll_x = ttk.Scrollbar(result_frame, orient='horizontal')
tree_scroll_x.pack(side="bottom", fill="x")

tree = ttk.Treeview(result_frame, yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
tree_scroll_y.config(command=tree.yview)
tree_scroll_x.config(command=tree.xview)
tree.pack(fill="both", expand=True)

def search():
    column_to_search = column_map[selected_field.get()]
    search_text = search_entry.get().strip().lower()

    if not search_text:
        messagebox.showwarning("""
                               Kush Toh Daalo
                               """
                               ,
                               """
                               Aare! Apne Koi Bhi Text Nahi Daala
                               Kush Toh Daalo Search Karne Ke Liye
                               Agar Koi Text Nahi Hai Toh Exit Karo!
                               Agar Hai Toh Search Karo!
                               
                               Dhanawaad!
                               
                               T_T
                               O_O
                               0_0
                               :-)
                               
                               """)
        return

    filtered = df[
        df[column_to_search].astype(str).str.strip().str.lower().str.contains(search_text, na=False)
    ]

    tree.delete(*tree.get_children())

    if filtered.empty:
        messagebox.showinfo("""
                            Galat Baat Ha
                            """
                            
                            ,
                            
                            """
                            Koi Data Exist Nahi Karta
                            Kush Galat Type Kiya ?
                            Agar Galat Type Kiya Toh Phir Se Try Karo!
                            Agar Nahi Kiya Toh Excel File Check Karo!
                            
                            Agar File Mein Data Hai
                            Lekin Error AA Raha Hai Toh
                            VPX Ko Contact Karo
                            
                            Dhanawaad!
                            Apka Error Solve Ho Jaye
                            
                            T_T
                            0_0
                            O_O
                            :-)
                            """)
        return

    tree["columns"] = list(filtered.columns)
    tree["show"] = "headings"
    for col in filtered.columns:
        tree.heading(col, text=col)
        tree.column(col, width=120, anchor="w")

    for _, row in filtered.iterrows():
        values = [str(val) for val in row]
        tree.insert("", "end", values=values)

tk.Button(root, text="Excel Se Data Dhundo", command=search, font=("Arial", 12), bg="#15FF00", fg="#000000", padx=20, pady=5).pack(pady=10)

root.mainloop()


t.sleep(0.1)
print(">>>")
t.sleep(0.1)
print(">>>")
t.sleep(0.1)
print(">>>")
t.sleep(0.1)
print(">>>")
t.sleep(0.1)
print(">>> Closing GUI And Software...")
t.sleep(0.5)
print(">>> Closed GUI And Software Successfully!")
t.sleep(0.5)
print(">>> Closing Excel File")
t.sleep(0.5)
print(">>> Closed Excel File Successfully!")
t.sleep(0.5)
print(">>> Thank you for using the Excel Search Software!")
t.sleep(0.5)
print(">>> This software was created by VPX")
t.sleep(0.5)
print(">>> How Was Your Experience?")
t.sleep(0.5)
Exit = input(">>> Press 'Enter' to Close ------> 'Enter' ").strip().lower()
print(">>> Exiting... Goodbye!")
t.sleep(0.5)
print(">>> Exit")
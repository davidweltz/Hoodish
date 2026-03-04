import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import os
import sys

# File Configuration
EXCEL_FILE = 'data_source.xlsx'
SHEET_NAME = 'Data'

def initialize_excel():
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=['Region', 'Department', 'Business Unit'])
        df.to_excel(EXCEL_FILE, index=False, sheet_name=SHEET_NAME)

def get_data():
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME).fillna('')
    except Exception:
        initialize_excel()
        return pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME).fillna('')

def save_data(df):
    df.to_excel(EXCEL_FILE, index=False, sheet_name=SHEET_NAME)

class App:
    def __init__(self, root):
        self.root = root
        # Instead of withdraw, we just keep the root as the main window 
        # to avoid "orphan" windows staying alive in the background.
        self.root.title("Business Unit Manager")
        self.root.geometry("800x650")
        
        # UI Styling
        self.bg_color = "#E0F2F7"   # Baby Blue
        self.context_fg = "#00008B"  # Dark Blue
        self.line_color = "#808080"  # Solid Gray
        
        self.root.configure(bg=self.bg_color)
        
        # Clean exit when clicking the 'X' on the main window
        self.root.protocol("WM_DELETE_WINDOW", self.safe_exit)
        
        self.build_main_frame()

    def safe_exit(self):
        """Kills the entire application and the background process."""
        self.root.quit()     # Stops the mainloop
        self.root.destroy()  # Destroys all widgets
        os._exit(0)         # Forces the system process to end immediately

    def refresh_visibility(self, condition, widgets, separator, window):
        if condition:
            separator.grid()
            for w in widgets: w.grid()
            window.columnconfigure(0, weight=1)
            window.columnconfigure(2, weight=1)
        else:
            separator.grid_remove()
            for w in widgets: w.grid_remove()
            window.columnconfigure(0, weight=1)
            window.columnconfigure(2, weight=0)

    # --- MAIN FRAME: REGION ---
    def build_main_frame(self):
        # We build directly on self.root now to prevent the "reappearing" loop
        for widget in self.root.winfo_children():
            widget.destroy()

        tk.Label(self.root, text="CREATE REGION", font=('Arial', 20, 'bold'), bg=self.bg_color, pady=25).grid(row=0, column=0, columnspan=3)

        self.sep_a = tk.Frame(self.root, width=5, bg=self.line_color)
        self.sep_a.grid(row=1, column=1, rowspan=4, sticky='ns', padx=20)

        # Left side: Creation
        tk.Label(self.root, text="Create Region Name", font=('Arial', 13, 'bold'), bg=self.bg_color).grid(row=1, column=0, pady=10)
        self.ent_region = tk.Entry(self.root, width=35, font=('Arial', 12), justify='center')
        self.ent_region.grid(row=2, column=0, padx=40, pady=5)
        
        tk.Button(self.root, text="Save, then Add Department", width=30, command=self.opt_a).grid(row=3, column=0, pady=8)
        tk.Button(self.root, text="Save", width=30, command=self.opt_b).grid(row=4, column=0, pady=8)

        # Right side: Modification
        self.lbl_mod_reg = tk.Label(self.root, text="Modify Region", font=('Arial', 13, 'bold'), bg=self.bg_color)
        self.cb_region = ttk.Combobox(self.root, width=32, state="readonly", font=('Arial', 11))
        self.btn_edit_reg = tk.Button(self.root, text="Edit Region Name", width=30, command=self.opt_c)
        self.btn_add_dept = tk.Button(self.root, text="Add Department", width=30, command=self.opt_d)

        self.reg_widgets = [self.lbl_mod_reg, self.cb_region, self.btn_edit_reg, self.btn_add_dept]
        for i, w in enumerate(self.reg_widgets):
            w.grid(row=i+1, column=2, padx=40, pady=5)

        tk.Button(self.root, text="Exit Program", width=18, bg="#ffcccb", command=self.safe_exit).grid(row=6, column=0, columnspan=3, pady=60)
        
        self.refresh_region_ui()

    def refresh_region_ui(self):
        df = get_data()
        regions = sorted([r for r in df['Region'].unique() if r != ''])
        self.cb_region['values'] = regions
        self.cb_region.set("Select Region...")
        self.refresh_visibility(len(regions) > 0, self.reg_widgets, self.sep_a, self.root)

    # ... (Rest of the functions like opt_a, show_frame_aa, etc. remain the same)
    # Just ensure every secondary window (Toplevel) also calls self.safe_exit if needed.

    def opt_a(self): 
        region = self.ent_region.get().strip()
        if not region:
            messagebox.showwarning("Warning", "Region Name cannot be blank")
            return
        df = get_data()
        if region in df['Region'].values:
            messagebox.showinfo("Exists", f"Region {region} already exists.")
            return
        new_row = pd.DataFrame([{'Region': region, 'Department': '', 'Business Unit': ''}])
        save_data(pd.concat([df, new_row], ignore_index=True))
        self.refresh_region_ui()
        self.show_frame_aa(region)

    def opt_b(self): 
        region = self.ent_region.get().strip()
        if not region:
            messagebox.showwarning("Warning", "Region Name cannot be blank")
            return
        df = get_data()
        if region in df['Region'].values:
            messagebox.showinfo("Exists", f"Region {region} already exists.")
        else:
            new_row = pd.DataFrame([{'Region': region, 'Department': '', 'Business Unit': ''}])
            save_data(pd.concat([df, new_row], ignore_index=True))
            messagebox.showinfo("Success", f"Region {region} added")
            self.ent_region.delete(0, tk.END)
        self.refresh_region_ui()

    def opt_c(self): 
        sel = self.cb_region.get()
        if sel and sel != "Select Region...": self.show_frame_ca(sel)

    def opt_d(self): 
        sel = self.cb_region.get()
        if not sel or sel == "Select Region...":
            messagebox.showwarning("Warning", "No Region Selected")
        else:
            self.show_frame_aa(sel)

    # --- FRAME CA: CHANGE REGION NAME ---
    def show_frame_ca(self, old_name):
        win_ca = tk.Toplevel(self.root)
        win_ca.title("Change Region Name")
        win_ca.configure(bg=self.bg_color)
        tk.Label(win_ca, text="CHANGE REGION NAME", font=('Arial', 14, 'bold'), bg=self.bg_color, pady=15).pack()
        tk.Label(win_ca, text=f"Edit Name for: {old_name}", bg=self.bg_color, font=('Arial', 11)).pack(pady=5)
        ent_new = tk.Entry(win_ca, width=30, font=('Arial', 12), justify='center')
        ent_new.insert(0, old_name)
        ent_new.pack(pady=10)

        def save_change():
            new_name = ent_new.get().strip()
            if not new_name:
                messagebox.showwarning("Warning", "Name cannot be blank")
                return
            df = get_data()
            df.loc[df['Region'] == old_name, 'Region'] = new_name
            save_data(df)
            win_ca.destroy()
            self.refresh_region_ui()

        tk.Button(win_ca, text="Save Changes", width=15, command=save_change).pack(side='left', padx=20, pady=20)
        tk.Button(win_ca, text="Cancel", width=15, command=win_ca.destroy).pack(side='right', padx=20, pady=20)

    # --- FRAME AA: CREATE DEPARTMENT ---
    def show_frame_aa(self, region):
        win_aa = tk.Toplevel(self.root)
        win_aa.title("Create Department")
        win_aa.geometry("800x600")
        win_aa.configure(bg=self.bg_color)

        tk.Label(win_aa, text="CREATE DEPARTMENT", font=('Arial', 20, 'bold'), bg=self.bg_color, pady=20).grid(row=0, columnspan=3)
        
        ctx_frame = tk.Frame(win_aa, bg=self.bg_color)
        ctx_frame.grid(row=1, columnspan=3, pady=10)
        tk.Label(ctx_frame, text="Region: ", font=('Arial', 14), fg=self.context_fg, bg=self.bg_color).pack(side='left')
        tk.Label(ctx_frame, text=region, font=('Arial', 14, 'bold'), fg=self.context_fg, bg=self.bg_color).pack(side='left')

        sep_aa = tk.Frame(win_aa, width=5, bg=self.line_color)
        sep_aa.grid(row=2, column=1, rowspan=4, sticky='ns', padx=20)

        tk.Label(win_aa, text="Create Department Name", font=('Arial', 13, 'bold'), bg=self.bg_color).grid(row=2, column=0)
        ent_dept = tk.Entry(win_aa, width=35, font=('Arial', 12), justify='center')
        ent_dept.grid(row=3, column=0, padx=40, pady=5)

        lbl_mod = tk.Label(win_aa, text="Modify Department", font=('Arial', 13, 'bold'), bg=self.bg_color)
        cb_dept = ttk.Combobox(win_aa, width=32, state="readonly", font=('Arial', 11))
        btn_edit = tk.Button(win_aa, text="Edit Department Name", command=lambda: self.show_frame_aca(region, cb_dept.get(), win_aa))
        btn_bu = tk.Button(win_aa, text="Add Business Unit", width=30, command=lambda: self.show_frame_aaa(region, cb_dept.get()) if cb_dept.get() != "Select Department..." else messagebox.showwarning("Warning", "No Department Selected"))

        dept_widgets = [lbl_mod, cb_dept, btn_edit, btn_bu]
        for i, w in enumerate(dept_widgets):
            w.grid(row=i+2, column=2, padx=40, pady=5)

        def refresh_dept_ui():
            df = get_data()
            depts = sorted([d for d in df[(df['Region'] == region) & (df['Department'] != '')]['Department'].unique()])
            cb_dept['values'] = depts
            cb_dept.set("Select Department...")
            self.refresh_visibility(len(depts) > 0, dept_widgets, sep_aa, win_aa)

        def save_dept(to_bu=False):
            d_name = ent_dept.get().strip()
            if not d_name:
                messagebox.showwarning("Warning", "Department Name cannot be blank")
                return
            df_curr = get_data()
            if ((df_curr['Region'] == region) & (df_curr['Department'] == d_name)).any():
                messagebox.showinfo("Exists", "Department already exists.")
                return
            mask = (df_curr['Region'] == region) & (df_curr['Department'] == '')
            if mask.any():
                df_curr.loc[mask.idxmax(), 'Department'] = d_name
            else:
                new_r = pd.DataFrame([{'Region': region, 'Department': d_name, 'Business Unit': ''}])
                df_curr = pd.concat([df_curr, new_r], ignore_index=True)
            save_data(df_curr)
            refresh_dept_ui()
            if to_bu: self.show_frame_aaa(region, d_name)
            ent_dept.delete(0, tk.END)

        tk.Button(win_aa, text="Save, then Add Business Unit", width=30, command=lambda: save_dept(True)).grid(row=4, column=0, pady=8)
        tk.Button(win_aa, text="Save", width=30, command=lambda: save_dept(False)).grid(row=5, column=0, pady=8)
        tk.Button(win_aa, text="Close Window", width=18, command=win_aa.destroy).grid(row=7, column=0, columnspan=3, pady=40)
        
        refresh_dept_ui()

    # --- FRAME ACA: CHANGE DEPARTMENT NAME ---
    def show_frame_aca(self, region, old_dept, parent_win):
        if old_dept == "Select Department..." or not old_dept: return
        win_aca = tk.Toplevel(parent_win)
        win_aca.title("Change Department Name")
        win_aca.configure(bg=self.bg_color)
        tk.Label(win_aca, text="CHANGE DEPARTMENT NAME", font=('Arial', 14, 'bold'), bg=self.bg_color, pady=15).pack()
        ent_new = tk.Entry(win_aca, width=30, font=('Arial', 12), justify='center')
        ent_new.insert(0, old_dept)
        ent_new.pack(pady=10)

        def save_aca():
            new_d = ent_new.get().strip()
            if not new_d: return
            df = get_data()
            df.loc[(df['Region'] == region) & (df['Department'] == old_dept), 'Department'] = new_d
            save_data(df)
            win_aca.destroy()

        tk.Button(win_aca, text="Save Changes", width=15, command=save_aca).pack(pady=20)

    # --- FRAME AAA: BUSINESS UNIT ---
    def show_frame_aaa(self, region, dept):
        if not dept or dept == "Select Department...": return
        win_aaa = tk.Toplevel(self.root)
        win_aaa.title("Create Business Unit")
        win_aaa.geometry("900x600")
        win_aaa.configure(bg=self.bg_color)

        tk.Label(win_aaa, text="CREATE BUSINESS UNIT", font=('Arial', 20, 'bold'), bg=self.bg_color, pady=20).grid(row=0, columnspan=3)
        
        ctx_frame = tk.Frame(win_aaa, bg=self.bg_color)
        ctx_frame.grid(row=1, columnspan=3, pady=10)
        tk.Label(ctx_frame, text="Region: ", font=('Arial', 14), fg=self.context_fg, bg=self.bg_color).pack(side='left')
        tk.Label(ctx_frame, text=region, font=('Arial', 14, 'bold'), fg=self.context_fg, bg=self.bg_color).pack(side='left')
        tk.Label(ctx_frame, text="  -----  Department: ", font=('Arial', 14), fg=self.context_fg, bg=self.bg_color).pack(side='left')
        tk.Label(ctx_frame, text=dept, font=('Arial', 14, 'bold'), fg=self.context_fg, bg=self.bg_color).pack(side='left')

        sep_aaa = tk.Frame(win_aaa, width=5, bg=self.line_color)
        sep_aaa.grid(row=2, column=1, rowspan=3, sticky='ns', padx=20)

        tk.Label(win_aaa, text="Create Business Unit Name", font=('Arial', 13, 'bold'), bg=self.bg_color).grid(row=2, column=0)
        ent_bu = tk.Entry(win_aaa, width=35, font=('Arial', 12), justify='center')
        ent_bu.grid(row=3, column=0, padx=40, pady=5)

        lbl_mod = tk.Label(win_aaa, text="Modify Business Unit", font=('Arial', 13, 'bold'), bg=self.bg_color)
        cb_bu = ttk.Combobox(win_aaa, width=32, state="readonly", font=('Arial', 11))
        btn_edit = tk.Button(win_aaa, text="Edit Business Unit Name", command=lambda: self.show_frame_aaba(region, dept, cb_bu.get(), win_aaa))
        
        bu_widgets = [lbl_mod, cb_bu, btn_edit]
        for i, w in enumerate(bu_widgets):
            w.grid(row=i+2, column=2, padx=40, pady=5)

        def refresh_bu_ui():
            df = get_data()
            bus = sorted([b for b in df[(df['Region']==region) & (df['Department']==dept) & (df['Business Unit']!='')]['Business Unit'].unique()])
            cb_bu['values'] = bus
            cb_bu.set("Select Business Unit...")
            self.refresh_visibility(len(bus) > 0, bu_widgets, sep_aaa, win_aaa)

        def save_bu():
            b_name = ent_bu.get().strip()
            if not b_name:
                messagebox.showwarning("Warning", "Business Unit cannot be blank")
                return
            df_curr = get_data()
            if ((df_curr['Region']==region) & (df_curr['Department']==dept) & (df_curr['Business Unit']==b_name)).any():
                messagebox.showinfo("Exists", "Business Unit already exists.")
                return
            mask = (df_curr['Region']==region) & (df_curr['Department']==dept) & (df_curr['Business Unit']=='')
            if mask.any():
                df_curr.loc[mask.idxmax(), 'Business Unit'] = b_name
            else:
                new_r = pd.DataFrame([{'Region': region, 'Department': dept, 'Business Unit': b_name}])
                df_curr = pd.concat([df_curr, new_r], ignore_index=True)
            save_data(df_curr)
            refresh_bu_ui()
            ent_bu.delete(0, tk.END)

        tk.Button(win_aaa, text="Save", width=30, command=save_bu).grid(row=4, column=0, pady=15)
        tk.Button(win_aaa, text="Close Window", width=18, command=win_aaa.destroy).grid(row=6, column=0, columnspan=3, pady=40)
        
        refresh_bu_ui()

    # --- FRAME AABA: EDIT BUSINESS UNIT ---
    def show_frame_aaba(self, region, dept, old_bu, parent):
        if old_bu == "Select Business Unit..." or not old_bu: return
        win_aaba = tk.Toplevel(parent)
        win_aaba.title("Edit Business Unit Name")
        win_aaba.configure(bg=self.bg_color)
        tk.Label(win_aaba, text="EDIT BUSINESS UNIT", font=('Arial', 14, 'bold'), bg=self.bg_color, pady=15).pack()
        ent_new = tk.Entry(win_aaba, width=30, font=('Arial', 12), justify='center')
        ent_new.insert(0, old_bu)
        ent_new.pack(pady=10)

        def save_aaba():
            new_b = ent_new.get().strip()
            if not new_b: return
            df = get_data()
            df.loc[(df['Region']==region) & (df['Department']==dept) & (df['Business Unit']==old_bu), 'Business Unit'] = new_b
            save_data(df)
            win_aaba.destroy()

        tk.Button(win_aaba, text="Save Changes", width=15, command=save_aaba).pack(pady=20)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
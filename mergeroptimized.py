import os
import pandas as pd
from tkinter import (
    Tk, Label, Entry, Button, Text, Scrollbar, StringVar, IntVar,
    END, filedialog, messagebox, font, OptionMenu, Checkbutton
)
from tkinter import ttk
from datetime import datetime


def read_excel_sheets(filepath):
    try:
        xl = pd.ExcelFile(filepath, engine='openpyxl')
        return xl.sheet_names
    except Exception:
        return []


def update_sheet_list(folder, master_filename, selected_sheet, sheet_menu):
    filepath = os.path.join(folder, master_filename)
    sheet_names = read_excel_sheets(filepath)
    sheet_menu['menu'].delete(0, 'end')
    if sheet_names:
        selected_sheet.set(sheet_names[0])
        for sheet in sheet_names:
            sheet_menu['menu'].add_command(label=sheet, command=lambda value=sheet: selected_sheet.set(value))
    else:
        selected_sheet.set("")
        messagebox.showwarning("Χωρίς φύλλα", f"Το αρχείο '{master_filename}' δεν περιέχει αναγνώσιμα φύλλα.")


def merge_excel_rows(folder, master_filename, output_filename, sheet_name, log, progress=None):
    merged_data = []
    failed_files = []
    success_count = 0
    output_path = os.path.join(folder, output_filename)

    def log_message(message):
        log.insert(END, message + "\n")
        log.see(END)

    try:
        master_path = os.path.join(folder, master_filename)
        master_df = pd.read_excel(master_path, sheet_name=sheet_name, engine='openpyxl', header=None)
        header = master_df.iloc[0].tolist()
        merged_data.append(header)
    except Exception as e:
        log_message(f"❌ Σφάλμα στο αρχείο master ή στο φύλλο '{sheet_name}': {e}")
        return

    excel_files = [f for f in os.listdir(folder) if f.endswith('.xlsx') and f not in [master_filename, output_filename]]

    for idx, filename in enumerate(excel_files):
        filepath = os.path.join(folder, filename)
        try:
            df = pd.read_excel(filepath, sheet_name=sheet_name, engine='openpyxl', header=None)
            if len(df) >= 2:
                rows_to_add = []
                for i in range(1, len(df)):
                    row = df.iloc[i]
                    if row.isnull().all() or all(str(cell).strip() == '' for cell in row):
                        break
                    rows_to_add.append(row.tolist())
                    log_message(f"✅ {filename} ➔ Γραμμή {i+1}: {row.tolist()}")
                if rows_to_add:
                    merged_data.extend(rows_to_add)
                    success_count += 1
                else:
                    failed_files.append((filename, "Η 2η γραμμή είναι εντελώς κενή ή δεν βρέθηκαν δεδομένα"))
            else:
                failed_files.append((filename, "Μόνο 1 γραμμή – δεν υπάρχει 2η για συγχώνευση"))
        except Exception as e:
            failed_files.append((filename, str(e)))

        if progress:
            progress['value'] = int(((idx + 1) / len(excel_files)) * 100)
            progress.update_idletasks()

    try:
        pd.DataFrame(merged_data).to_excel(output_path, index=False, header=False, engine='openpyxl')
        log_message(f"📂 Το αρχείο συγχωνεύτηκε με επιτυχία: {output_filename}")
    except Exception as e:
        log_message(f"❌ Σφάλμα κατά την αποθήκευση του αρχείου: {e}")
        return

    log_message(f"\n📊 Συνολικά αρχεία: {len(excel_files)}")
    log_message(f"✅ Επιτυχώς συγχωνεύθηκαν: {success_count}")
    log_message(f"⚠ Προβληματικά αρχεία: {len(failed_files)}")
    for f, reason in failed_files:
        log_message(f"  - {f}: {reason}")


def save_log_to_file(folder_path, log_widget):
    log_content = log_widget.get("1.0", END).strip()
    if not log_content:
        return
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_filename = f"merge_log_{timestamp}.txt"
    log_path = os.path.join(folder_path, log_filename)
    try:
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write(log_content)
        log_widget.insert(END, f"\n📝 Το log αποθηκεύτηκε στο αρχείο: {log_filename}\n")
    except Exception as e:
        log_widget.insert(END, f"\n❌ Σφάλμα κατά την αποθήκευση του log: {e}\n")


def main():
    def start_merge():
        folder = folder_entry.get()
        master = master_entry.get()
        output = output_entry.get()
        sheet = selected_sheet.get()
        output_path = os.path.join(folder, output)

        if not os.path.isdir(folder):
            messagebox.showerror("Σφάλμα", "Ο φάκελος δεν υπάρχει.")
            return
        if not os.path.exists(os.path.join(folder, master)):
            messagebox.showerror("Σφάλμα", "Το αρχείο master δεν βρέθηκε.")
            return

        if os.path.exists(output_path):
            if not messagebox.askyesno("Υπάρχει ήδη αρχείο", f"Το αρχείο '{output}' υπάρχει ήδη. Θέλεις να διαγραφεί;"):
                log_text.insert(END, "ℹ️ Η διαδικασία ακυρώθηκε από τον χρήστη.\n")
                return
            try:
                os.remove(output_path)
            except Exception as e:
                messagebox.showerror("Σφάλμα διαγραφής", f"Δεν ήταν δυνατή η διαγραφή του αρχείου: {e}")
                return

        log_text.delete('1.0', END)
        progress_bar['value'] = 0
        merge_excel_rows(folder, master, output, sheet, log_text, progress=progress_bar)
        save_log_to_file(folder, log_text)

    def browse_folder():
        path = filedialog.askdirectory()
        if path:
            folder_entry.delete(0, END)
            folder_entry.insert(0, path)
            update_sheet_list(path, master_entry.get(), selected_sheet, sheet_menu)

    def browse_master_file():
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            folder, filename = os.path.split(path)
            folder_entry.delete(0, END)
            folder_entry.insert(0, folder)
            master_entry.delete(0, END)
            master_entry.insert(0, filename)
            update_sheet_list(folder, filename, selected_sheet, sheet_menu)

    def master_changed(*args):
        update_sheet_list(folder_entry.get(), master_entry.get(), selected_sheet, sheet_menu)

    def toggle_dark_mode():
        if dark_mode_var.get():
            window.configure(bg="#2e2e2e")
            log_text.configure(bg="#1e1e1e", fg="white")
            for widget in window.winfo_children():
                if isinstance(widget, (Label, Button, Entry, OptionMenu)):
                    widget.configure(bg="#2e2e2e", fg="white")
        else:
            window.configure(bg="SystemButtonFace")
            log_text.configure(bg="white", fg="black")
            for widget in window.winfo_children():
                if isinstance(widget, (Label, Button, Entry, OptionMenu)):
                    widget.configure(bg="SystemButtonFace", fg="black")

    def close_app():
        window.destroy()

    window = Tk()
    dark_mode_var = IntVar()
    window.title("Συγχώνευση Excel αρχείων")
    window.geometry("1000x700")
    window.minsize(800, 550)

    window.grid_rowconfigure(5, weight=1)
    window.grid_columnconfigure(1, weight=1)

    label_font = font.Font(size=11, weight='bold')
    button_font = font.Font(size=11)
    entry_font = font.Font(size=10)
    log_font = font.Font(family="Courier", size=9)

    Label(window, text="📂 Φάκελος:", font=label_font).grid(row=0, column=0, sticky='e')
    folder_entry = Entry(window, width=60, font=entry_font)
    folder_entry.insert(0, "merge_files")
    folder_entry.grid(row=0, column=1, padx=5, pady=3, sticky='ew')
    Button(window, text="Επιλογή...", font=button_font, command=browse_folder).grid(row=0, column=2, padx=5)

    Label(window, text="📄 master αρχείο:", font=label_font).grid(row=1, column=0, sticky='e')
    master_entry = Entry(window, width=60, font=entry_font)
    master_entry.insert(0, "master.xlsx")
    master_entry.grid(row=1, column=1, padx=5, pady=3, sticky='ew')
    master_entry.bind("<FocusOut>", master_changed)
    Button(window, text="Επιλογή...", font=button_font, command=browse_master_file).grid(row=1, column=2, padx=5)

    Label(window, text="📝 Αρχείο εξόδου:", font=label_font).grid(row=2, column=0, sticky='e')
    output_entry = Entry(window, width=60, font=entry_font)
    output_entry.insert(0, "merged_output.xlsx")
    output_entry.grid(row=2, column=1, padx=5, pady=3, sticky='ew')

    Label(window, text="📑 Επιλογή φύλλου:", font=label_font).grid(row=3, column=0, sticky='e')
    selected_sheet = StringVar()
    sheet_menu = OptionMenu(window, selected_sheet, "")
    sheet_menu.grid(row=3, column=1, padx=5, pady=3, sticky='ew')
    Button(window, text="🔄 Ανάγνωση φύλλων", font=button_font, command=lambda: update_sheet_list(folder_entry.get(), master_entry.get(), selected_sheet, sheet_menu)).grid(row=3, column=2)

    Button(window, text="🚀 Έναρξη συγχώνευσης", font=button_font, command=start_merge).grid(row=4, column=1, pady=10)

    progress_bar = ttk.Progressbar(window, orient="horizontal", length=400, mode="determinate")
    progress_bar.grid(row=5, column=1, pady=5)

    log_text = Text(window, font=log_font)
    log_text.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky='nsew')

    scrollbar = Scrollbar(window, command=log_text.yview)
    log_text.configure(yscrollcommand=scrollbar.set)
    scrollbar.grid(row=6, column=3, sticky='ns')

    Checkbutton(window, text="🌙 Dark Mode", variable=dark_mode_var, command=toggle_dark_mode, font=button_font).grid(row=7, column=0, pady=5, sticky='w')
    Button(window, text="❌ Κλείσιμο", font=button_font, command=close_app).grid(row=7, column=1, pady=5)

    update_sheet_list("merge_files", "master.xlsx", selected_sheet, sheet_menu)

    window.mainloop()


if __name__ == "__main__":
    main()
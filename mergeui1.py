import os
import pandas as pd
from tkinter import Tk, Label, Entry, Button, Text, filedialog, messagebox, Scrollbar, END, font
from datetime import datetime

def merge_excel_rows(folder_path, master_filename, output_filename, sheet_name, log_widget):
    def log(message):
        log_widget.insert(END, message + '\n')
        log_widget.see(END)
        log_widget.update()

    master_path = os.path.join(folder_path, master_filename)
    output_path = os.path.join(folder_path, output_filename)
    log_lines = []
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_lines.append(f"=== Αναφορά συγχώνευσης αρχείων Excel ({timestamp}) ===\n")

    try:
        master_df = pd.read_excel(master_path, header=None, nrows=1, sheet_name=sheet_name, engine='openpyxl')
    except Exception as e:
        msg = f"❌ Σφάλμα στο master αρχείο: {e}"
        log(msg)
        return

    merged_data = [master_df.iloc[0].tolist()]
    total_files = 0
    success_count = 0
    failed_files = []

    for filename in os.listdir(folder_path):
        if not filename.endswith('.xlsx') or filename in [master_filename, output_filename]:
            continue

        total_files += 1
        file_path = os.path.join(folder_path, filename)

        try:
            df = pd.read_excel(file_path, header=None, sheet_name=sheet_name, engine='openpyxl')

            if len(df) >= 2:
                row_data = df.iloc[1].tolist()
                merged_data.append(row_data)
                success_count += 1
                log(f"✅ {filename}")
                log(f"   ➤ 2η γραμμή: {row_data}")
            else:
                failed_files.append((filename, "Λιγότερες από 2 γραμμές"))
                log(f"⚠ {filename}: Λιγότερες από 2 γραμμές")

        except ValueError:
            failed_files.append((filename, "Φύλλο δεν βρέθηκε"))
            log(f"⚠ {filename}: Φύλλο '{sheet_name}' δεν βρέθηκε")
        except Exception as e:
            failed_files.append((filename, str(e)))
            log(f"⚠ {filename}: {e}")

    try:
        merged_df = pd.DataFrame(merged_data)
        merged_df.to_excel(output_path, header=False, index=False)
        log(f"\n✅ Το αρχείο εξόδου δημιουργήθηκε: {output_path}")
    except Exception as e:
        log(f"❌ Σφάλμα αποθήκευσης εξόδου: {e}")
        return

    log("\n📊 Αναφορά:")
    log(f"🔢 Συνολικά αρχεία (εκτός master): {total_files}")
    log(f"✅ Επιτυχώς διαβασμένα: {success_count}")
    log(f"❌ Προβληματικά: {len(failed_files)}")
    for fname, reason in failed_files:
        log(f" - {fname}: {reason}")

def save_log_to_file(folder_path, log_widget):
    # Παίρνει όλα τα περιεχόμενα του Text widget
    log_content = log_widget.get("1.0", END).strip()

    if not log_content:
        return  # Αν το log είναι κενό, δεν κάνουμε τίποτα

    # Δημιουργία αρχείου καταγραφής με timestamp
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
        sheet = sheet_entry.get()
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
        merge_excel_rows(folder, master, output, sheet, log_text)
        save_log_to_file(folder, log_text)

    def browse_folder():
        path = filedialog.askdirectory()
        if path:
            folder_entry.delete(0, END)
            folder_entry.insert(0, path)

    def close_app():
        window.destroy()

    # --- GUI SETUP ---
    window = Tk()
    window.title("Συγχώνευση Excel αρχείων")
    window.geometry("1000x650")
    window.minsize(800, 500)

    # Grid δυναμικότητα
    window.grid_rowconfigure(5, weight=1)
    window.grid_columnconfigure(1, weight=1)

    # Fonts
    label_font = font.Font(size=11, weight='bold')
    entry_font = font.Font(size=10)
    log_font = font.Font(family="Courier", size=9)

    # Ετικέτες και πεδία
    Label(window, text="📂 Φάκελος:", font=label_font).grid(row=0, column=0, sticky='e')
    folder_entry = Entry(window, width=60, font=entry_font)
    folder_entry.insert(0, "merge_files")
    folder_entry.grid(row=0, column=1, padx=5, pady=3, sticky='ew')
    Button(window, text="Επιλογή...", command=browse_folder).grid(row=0, column=2, padx=5)

    Label(window, text="📄 master αρχείο:", font=label_font).grid(row=1, column=0, sticky='e')
    master_entry = Entry(window, width=60, font=entry_font)
    master_entry.insert(0, "master.xlsx")
    master_entry.grid(row=1, column=1, padx=5, pady=3, sticky='ew')

    Label(window, text="📝 Αρχείο εξόδου:", font=label_font).grid(row=2, column=0, sticky='e')
    output_entry = Entry(window, width=60, font=entry_font)
    output_entry.insert(0, "merged_output.xlsx")
    output_entry.grid(row=2, column=1, padx=5, pady=3, sticky='ew')

    Label(window, text="📑 Φύλλο Excel:", font=label_font).grid(row=3, column=0, sticky='e')
    sheet_entry = Entry(window, width=60, font=entry_font)
    sheet_entry.grid(row=3, column=1, padx=5, pady=3, sticky='ew')

    Button(window, text="🚀 Έναρξη συγχώνευσης", command=start_merge).grid(row=4, column=1, pady=10)

    # Περιοχή εμφάνισης αποτελεσμάτων (log)
    log_text = Text(window, font=log_font)
    log_text.grid(row=5, column=0, columnspan=3, padx=10, pady=10, sticky='nsew')

    scrollbar = Scrollbar(window, command=log_text.yview)
    log_text.configure(yscrollcommand=scrollbar.set)
    scrollbar.grid(row=5, column=3, sticky='ns')

    # Κουμπί Κλείσιμο
    Button(window, text="❌ Κλείσιμο", command=close_app).grid(row=6, column=1, pady=5)

    window.mainloop()

if __name__ == '__main__':
    main()

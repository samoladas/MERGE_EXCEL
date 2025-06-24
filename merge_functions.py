import os
import pandas as pd
from datetime import datetime

def merge_excel_rows(folder_path, master_filename, output_filename, sheet_name):
    master_path = os.path.join(folder_path, master_filename)
    log_lines = []
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_lines.append(f"=== Αναφορά συγχώνευσης αρχείων Excel ({timestamp}) ===\n")

    try:
        master_df = pd.read_excel(master_path, header=None, nrows=1, sheet_name=sheet_name, engine='openpyxl')
    except Exception as e:
        error_msg = f"❌ Σφάλμα στο master αρχείο '{master_filename}' (φύλλο '{sheet_name}'): {e}"
        print(error_msg)
        log_lines.append(error_msg)
        write_log(folder_path, log_lines)
        return

    merged_data = [master_df.iloc[0].tolist()]
    total_files = 0
    success_count = 0
    failed_files = []

    for filename in os.listdir(folder_path):
        if not filename.endswith('.xlsx') or filename == master_filename:
            continue

        total_files += 1
        file_path = os.path.join(folder_path, filename)

        try:
            df = pd.read_excel(file_path, header=None, sheet_name=sheet_name, engine='openpyxl')

            if len(df) >= 2:
                row_data = df.iloc[1].tolist()
                merged_data.append(row_data)
                success_count += 1

                msg = f"✅ Επεξεργασία αρχείου: {filename}"
                row_msg = f"   ➤ 2η γραμμή: {row_data}"
                print(msg)
                print(row_msg)
                log_lines.append(msg)
                log_lines.append(row_msg)
            else:
                reason = "Λιγότερες από 2 γραμμές"
                failed_files.append((filename, reason))
                msg = f"⚠ Το αρχείο '{filename}' δεν έχει τουλάχιστον 2 γραμμές. Παραλείπεται."
                print(msg)
                log_lines.append(msg)

        except ValueError:
            reason = "Φύλλο δεν βρέθηκε"
            failed_files.append((filename, reason))
            msg = f"⚠ Το φύλλο '{sheet_name}' δεν βρέθηκε στο αρχείο '{filename}'. Παραλείπεται."
            print(msg)
            log_lines.append(msg)
        except Exception as e:
            reason = str(e)
            failed_files.append((filename, reason))
            msg = f"⚠ Σφάλμα στο αρχείο '{filename}': {e}. Παραλείπεται."
            print(msg)
            log_lines.append(msg)

    # Αποθήκευση αποτελεσμάτων
    output_path = os.path.join(folder_path, output_filename)
    try:
        merged_df = pd.DataFrame(merged_data)
        merged_df.to_excel(output_path, header=False, index=False)
        success_msg = f"\n✅ Το αρχείο εξόδου δημιουργήθηκε επιτυχώς: {output_path}"
        print(success_msg)
        log_lines.append(success_msg)
    except Exception as e:
        error_msg = f"❌ Σφάλμα κατά την αποθήκευση του αρχείου εξόδου: {e}"
        print(error_msg)
        log_lines.append(error_msg)

    # Τελική αναφορά
    report = [
        "\n📊 Αναφορά:",
        f"🔢 Συνολικά αρχεία (εκτός master): {total_files}",
        f"✅ Επιτυχώς διαβασμένα: {success_count}",
        f"❌ Προβληματικά αρχεία: {len(failed_files)}"
    ]
    print('\n'.join(report))
    log_lines.extend(report)

    if failed_files:
        log_lines.append("📌 Λίστα προβληματικών:")
        for fname, reason in failed_files:
            line = f" - {fname}: {reason}"
            print(line)
            log_lines.append(line)

    write_log(folder_path, log_lines)

def write_log(folder_path, lines):
    log_path = os.path.join(folder_path, 'merge_log.txt')
    try:
        with open(log_path, 'w', encoding='utf-8') as f:
            for line in lines:
                f.write(line + '\n')
        print(f"\n📝 Αρχείο καταγραφής δημιουργήθηκε: {log_path}")
    except Exception as e:
        print(f"❌ Σφάλμα κατά την αποθήκευση του αρχείου log: {e}")

def main():
    folder_path = 'merge_files'
    master_filename = 'master.xlsx'
    output_filename = 'merged_output.xlsx'
    sheet_name = 'Δημοτικά'
    
    output_path = os.path.join(folder_path, output_filename)

    # 🔁 Έλεγχος αν υπάρχει ήδη το αρχείο εξόδου
    if os.path.exists(output_path):
        answer = input(f"❓ Το αρχείο '{output_filename}' υπάρχει ήδη. Θέλεις να διαγραφεί; (ν/ο): ").strip().lower()
        if answer == 'ν':
            try:
                os.remove(output_path)
                print(f"🗑️ Το αρχείο '{output_filename}' διαγράφηκε.")
            except Exception as e:
                print(f"❌ Σφάλμα κατά τη διαγραφή του αρχείου: {e}")
        else:
            print("ℹ️ Το αρχείο δεν διαγράφηκε. Ενδέχεται να γίνει αντικατάσταση εάν δημιουργηθεί νέο.")

    merge_excel_rows(folder_path, master_filename, output_filename, sheet_name)

if __name__ == '__main__':
    main()

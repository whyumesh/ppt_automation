import os
import shutil

SOURCE = r"C:\PATH\TO\YOUR\FOLDER"      # 🔴 CHANGE THIS
DEST = r"E:\company_backup"             # 🔴 CHANGE THIS (USB path)

def copy_with_resume(src, dst):
    total_files = 0
    copied_files = 0

    for root, dirs, files in os.walk(src):
        total_files += len(files)

    print(f"Total files to process: {total_files}\n")

    for root, dirs, files in os.walk(src):
        rel_path = os.path.relpath(root, src)
        dest_dir = os.path.join(dst, rel_path)
        os.makedirs(dest_dir, exist_ok=True)

        for file in files:
            src_file = os.path.join(root, file)
            dst_file = os.path.join(dest_dir, file)

            if os.path.exists(dst_file):
                continue

            try:
                shutil.copy2(src_file, dst_file)
                copied_files += 1
                print(f"[{copied_files}/{total_files}] Copied: {file}")
            except Exception as e:
                print(f"❌ Failed: {src_file} → {e}")

    print("\n✅ COPY COMPLETED SUCCESSFULLY")

if __name__ == "__main__":
    copy_with_resume(SOURCE, DEST)
    input("\nPress Enter to exit...")

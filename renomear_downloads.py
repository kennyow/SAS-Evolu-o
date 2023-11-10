import os
import openpyxl

download_directory = r"C:\Users\citee\Downloads"
new_names = ["new_name1.xls", "new_name2.xls", "new_name3.xls"]

# List all files in the download directory with their timestamps
files = [(filename, os.path.getmtime(os.path.join(download_directory, filename))) for filename in os.listdir(download_directory)]

# Sort the files by their timestamps in descending order (most recent first)
files.sort(key=lambda x: x[1], reverse=True)

# Rename the last 3 downloaded files
for i, (filename, _) in enumerate(files[:3]):
    if i < len(new_names):
        new_name = new_names[i]
        old_path = os.path.join(download_directory, filename)
        new_path = os.path.join(download_directory, new_name)
        os.rename(old_path, new_path)
        print(f"Renamed '{filename}' to '{new_name}'")

if len(files) == 0:
    print("No files found in the download directory.")
elif len(files) < 3:
    print("Less than 3 files found in the download directory. Renamed as many as available.")



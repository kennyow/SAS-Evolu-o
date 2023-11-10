import os


# Directory where you download files
download_directory = r"C:\Users\citee\Downloads"

# List files in the download directory
files = os.listdir(download_directory)

# Filter out directories and keep only files
files = [f for f in files if os.path.isfile(os.path.join(download_directory, f))]

# Sort files by modification time in descending order (most recent first)
files.sort(key=lambda x: os.path.getmtime(os.path.join(download_directory, x)), reverse=True)

# Check if there are any downloaded files
if len(files) > 0:
    # Get the name of the most recently downloaded file
    last_downloaded_file2 = files[0]
    input_string = str(last_downloaded_file2[:-4])

    # Print and save the name of the last downloaded file
    print("Last Downloaded File:", input_string)

else:
    print("No downloaded files found in the directory.")


# Split the string by underscore "_"
split_result = input_string.split("_")

pros= str(split_result[4])
if len(pros) > 2:
    pros = str(split_result[5])
avaliacao = str(split_result[1]).upper()
ano = str(pros[:-1])

# Print the result
print("Split Result:", split_result)
print("Year:", ano)
print("Avaliação:", avaliacao)
print(len(avaliacao))

# Construct the path using the year
original_path = r'G:\Drives compartilhados\_Anos Finais\Anos Finais - 2023\Notas de Avaliações\XXº Ano\II TRIMESTRE\XXº ano - WW - II TRIMESTRE'

# Replace '8°' with '9°' in the path
new_path2 = original_path.replace('XX', ano)
new_path = new_path2.replace('WW', avaliacao)

print(new_path)

try:
    os.startfile(new_path)
except Exception as e:
    print(f"An error occurred: {e}")
import urllib.request
import zipfile
import os
import sys
import ctypes
import getpass

# Check if running with administrator rights
#if not sys.platform.startswith('win') or not ctypes.windll.shell32.IsUserAnAdmin():
   # print("Error: This program requires administrator rights to run.")
   # input("Press Enter to exit...")
   # sys.exit(1)

# Check internet connectivity
try:
    urllib.request.urlopen('http://www.google.com', timeout=2)
except urllib.error.URLError:
    print("Error: No internet connection.")
    input("Press Enter to exit...")
    sys.exit(1)

print("10%: Internet connection established")


# URL of the zip file
#url 1 is excel tools
url1 = ''
#url 2 is excel file
url2 = ''

# Directory to extract the zip files
extract_dir1 = 'C:/LHITS Tools'
extract_dir2 = f'C:/Users/{getpass.getuser()}/AppData/Local/Microsoft/Office'

# Create the extraction directories if they don't exist
try:
    if not os.path.exists(extract_dir1):
        os.makedirs(extract_dir1)
    #if not os.path.exists(extract_dir2):
    #    os.makedirs(extract_dir2)
except OSError:
    print("Error: Failed to create the directory.")
    input("Press Enter to exit...")
    sys.exit(1)

print("20%: Directories found or created")

# Download the first zip file
try:
    filename1, _ = urllib.request.urlretrieve(url1)
except urllib.error.URLError:
    print("Error: Failed to download the LHITS Tools zip file.")
    input("Press Enter to exit...")
    sys.exit(1)
    
print("40%: Downloaded the tools successfully")

# Extract the first zip file
try:
    with zipfile.ZipFile(filename1, 'r') as zip_ref:
        zip_ref.extractall(extract_dir1)
except zipfile.BadZipFile:
    print("Error: Failed to extract the LHITS Tools zip file.")
    input("Press Enter to exit...")
    sys.exit(1)
    
print("60%: Extracted the tools successfully")

# Remove the downloaded first zip file
os.remove(filename1)

# Download the second zip file
try:
    filename2, _ = urllib.request.urlretrieve(url2)
except urllib.error.URLError:
    print("Error: Failed to download the Excel UI zip file.")
    input("Press Enter to exit...")
    sys.exit(1)
    
print("70%: Downloaded the excel ribbon file successfully")

# Extract the second zip file
try:
    with zipfile.ZipFile(filename2, 'r') as zip_ref:
        # Delete any existing files with the same name
        for file in zip_ref.namelist():
            file_path = os.path.join(extract_dir2, file)
            if os.path.exists(file_path):
                os.remove(file_path)
        zip_ref.extractall(extract_dir2)
except zipfile.BadZipFile:
    print("Error: Failed to extract the Excel UI zip file.")
    input("Press Enter to exit...")
    sys.exit(1)
    
print("90%: Installed the excel ribbon file successfully")

# Remove the downloaded second zip file
os.remove(filename2)

print("100% -----------------------------------------------------------------------------------------------")
print("----------------------------------------------------------------------------------------------------")
print('LHTIS Tools downloaded and installed successfully.')
print('Next time you open Excel you will see the new buttons on your ribbon.')
print("----------------------------------------------------------------------------------------------------")
input("Press Enter to exit...")

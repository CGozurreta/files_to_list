# files_to_list
This is a very simple Python applications that takes the file names in folders and turns them into bullet points in an exported Word file

Requirements:
- Python 3.x
- tkinter
- docx

Instructions:

After running the script, you will be presented with a window prompting you to pick the root folder you want to use (make sure it's the root folder that holds all your folders and files)

After selecting the right folder, you will be asked two questions:
- Do you want to keep full names? - The output text will display the full name, including file extensions
- Do you want to keep _? - The _ characters will be replaced with spaces (except for folder names)

After the two options were selected, the script will create a file called "items in folders list" in the folder where the pythin script is located
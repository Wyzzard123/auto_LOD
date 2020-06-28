# README

Creates an automatic (unformatted) list of documents based on a naming pattern.

## How to Use:

1. Download Python here: 
https://www.python.org/downloads/

2. Download these files by opening a terminal (eg Powershell or Bash) and type the following and press enter:
    ``` 
    git clone https://github.com/Wyzzard123/auto_LOD.git
    ```

3. There should now be a folder called auto_LOD. Either open a terminal from that folder or type the following to change your directory to it:
    ``` 
   cd auto_LOD 
   ```

4. To install all the requirements needed, type the following:
    ```
    pip install -r requirements.txt
    ```
   
5. If you had problems with installing the requirements.txt, try the following to upgrade your package manager first before installing the requirements:
    ``` 
    easy_install -U pip
    pip install -r requirements.txt
    ```
   
1. Open up auto_LOD.py in any text editor (like Notepad or VS Code) or IDE.

1. Change the file name that you want in the source code 
   ```
   word_doc_file_name = 'New List of Documents.docx'
   ```

6. Change the lod_files_directory variable in the source code to the folder where everything is. For example:
    ``` 
   lod_files_directory = 'C:/Documents/Case/Named Documents' 
   ```

7. Make sure all the files in the directory are named as follows:
    
   - <YYYY.MM.DD> <optional: HH.MM> <Document Name>.<file extension>
   - e.g. 2020.01.12 23.59 abcd.pdf, 2020.05.22 abcd.docx 

7. Then run the script in the terminal:

    ```
    python auto_LOD.py
    ```

You should have an automatically made List of Documents with the filename/at the file path that you provided.

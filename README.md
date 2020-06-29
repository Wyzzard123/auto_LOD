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
   (__OPTIONAL:__ You may also use a virtual environment if you know how)
   
5. If you had problems with installing the requirements.txt, try the following to upgrade your package manager first before installing the requirements:
    ``` 
    easy_install -U pip
    pip install -r requirements.txt
    ```

7. Make sure all the files in the directory are named as follows:
    
   - <YYYY.MM.DD> <optional: HH.MM> <Document Name>.<file extension>
   - e.g. 2020.01.12 23.59 abcd.pdf, 2020.05.22 abcd.docx 

7. Then run the script in the terminal. You may add a directory (-d or --directory) and a filename (-n --name). By default, these
will be set to 'test_files' and 'New LOD.docx' respectively. 

    ``` 
   python auto_LOD.py [-d path to directory containing LOD files] [-n file name or path to new list of documents]
    ```
   You can replace the directory name and file names in the example below:
   ``` 
   python auto_LOD.py -d 'C:\Users\user\Documents\Client Files\All Documents Compiled' -n 'C:\Users\user\Documents\List of Documents dated xx.xx.docx'
   ```
   If you do not provide a file path and instead only provide a file name, the file will appear in the same folder as the auto_LOD.py file.
   Use the quotation marks to avoid any issues with spaces.
   
8. If there are any files which do not follow the required naming pattern, you will be prompted to Continue (Y) or Abort (N). 
    * Type y to continue and n to abort. 
    * At the very end, you will be prompted one more time and presented with a list of failed names. 
    * If you continue, the program will create the list of documents without the failed files.
    * Otherwise, the program will not continue, and you can rename your files accordingly.
    * You can also keep continuing in order to get a complete list of invalid names. 

9. You should now have an automatically made List of Documents with the filename/at the file path that you provided.

## NOTES

1. If the folder provided in the directory_name argument does not exist, the program will exit.

1. If the path to the file name provided does not exist, the program will create intermediate folders.

1. If you wish to categorize the List of Documents by means other than date, you can put the documents into different folders and run the script on each folder, then compile the results together.

1. You will have to provide your own formatting.

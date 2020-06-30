# README

Creates an automatic (unformatted) list of documents based on a naming pattern.

## How to Use:

1. Download Python here: 
https://www.python.org/downloads/

2. Download these files by opening a terminal (eg Powershell or Bash) and type the following and press enter:
    ``` 
    git clone https://github.com/Wyzzard123/auto_LOD.git
    ```
   __NOTE__: To find a terminal on windows or Mac:
   * For Windows: 
       * Type Powershell in the Windows Search bar and open it. 
       * Alternatively, to get into the relevant folder, you can go to your file explorer and head to the relevant folder.
       * Then hold shift and right click an empty space and click "Open PowerShell Window here"
   hold shift, then right click an empty space in the folder and click
   * For Mac:
       * Click the Launchpad icon in the Dock 
       * Type Terminal in the search field. 
       * Click Terminal.

3. There should now be a folder called auto_LOD. Either open a terminal from that folder or type the following to change your directory to it:
    ``` 
   cd auto_LOD 
   ```
   __NOTE__: Alternatively, go to the new auto_LOD folder in your file explorer and open the document there.

4. To install all the requirements needed, type the following into the terminal (making sure you are in the auto_LOD folder) and press enter:
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
    
    The format is as follows: 

    ``` 
   python auto_LOD.py [-d path to directory containing LOD files] [-n file name or path to new list of documents]
    ```
   Alternatively, type the following into your terminal window, then replace the directory name and file names in the quotes:
   ``` 
   python auto_LOD.py -d 'C:\Users\user\Documents\Client Files\All Documents Compiled' -n 'C:\Users\user\Documents\List of Documents dated xx.xx.docx'
   ```
   If you do not provide a file path and instead only provide a file name, the file will appear in the same folder as the auto_LOD.py file.
   Use the quotation marks to avoid any issues with spaces.
   
   __NOTE__: You generally cannot click to navigate in a terminal window, and will have to use the arrow keys to move back and forth and delete text.
   
8. If there are any files which do not follow the required naming pattern, you will be prompted to Continue/Skip the invalidly named files (Y) or Abort/Exit (N). 
    * Type 'y' to continue and 'n' to abort then press enter. 
    * At the very end, you will be prompted one more time and presented with a list of failed names. 
    * If you continue, the program will create the list of documents without the failed files.
    * Otherwise, the program will not continue, and you can rename your files accordingly.
    * You can also keep continuing in order to get a complete list of invalid names. 

9. You should now have an automatically made List of Documents with the filename/at the file path that you provided.

10. To get updates and make sure you have the latest version, type the following into the terminal while in the auto_LOD folder:

    ``` 
    git pull origin master
    ```

## NOTES

1. If the folder provided in the directory_name argument does not exist, the program will exit.

1. If the path to the file name provided does not exist, the program will create intermediate folders.

1. If you wish to categorize the List of Documents by means other than date, you can put the documents into different folders and run the script on each folder, then compile the results together.

1. You will have to provide your own formatting.

1. IMPORTANT: The author of the document will be listed as python-docx. Change this on your own.

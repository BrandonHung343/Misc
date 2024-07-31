How to Set Up for Other Computers

1) After cloning this repo from Git into the proper file (contact Brandon for help), open up a command prompt window
2) In the cmd prompt window, navigate to the repo
3) Once in the repo, type the line "pip install -r requirements.txt"
4) Once the libraries have been installed, run the included tesseract.exe file 
5) Copy the files from Z:\Shipping\zzINVOICES & DOX\Invoice_Sorter\imagemagick into C:\Program Files\ImageMagick-6.9.10-Q16
6) Open up the Windows search and type "path". Click on "Edit the system environment variables".
7) In the bottom window labelled "System Variables", click Path -> Edit -> Browse -> Look for C:\Program Files\ImageMagick-6.9.10-Q16 -> Ok
8) In the bottom window labelled "System Variables", click New. Set Variable name to MAGICK_HOME and variable value to C:\Program Files\ImageMagick-6.9.10-Q16   
9) In the bottom window labelled "System Variables", click New. Set Variable name to MAGICK_HOME and variable value to C:\Program Files\ImageMagick-6.9.10-Q16\modules\coders
10) Open the command prompt. Type in "python" and press enter. 
11) Once python opens (next to the cursor will be a >>>), type "import pyocr" and press enter.
12) Type "print(pyocr.__file__)" and press enter. The file location will print for you. 
13) Navigate to the location you found in step 12 on the File Explorer. Right click on builders.py, click "Edit with IDLE"
14) In IDLE, use Edit -> Replace and replace every instance of "-psm" with "--psm". There should be around three of them.
14b) You might need to install Ghostfire as well. 
15) Save the file and exit. You're all set to run the invoice sorter now!



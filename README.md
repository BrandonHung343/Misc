# How to Set Up for Other Computers
## This stack is meant to run on Python 3.6. It has not been tested for other versions of Python.
1) After cloning this repo from Git into the proper file, open up a command prompt window
2) In the cmd prompt window, navigate to this repo
3) Once in the repo, type the line "pip install -r requirements.txt"
4) Once the libraries have been installed, extract the `tesseract.zip` file and run the extracted `tesseract.exe` file 
5) Install Imagemagick and add it to path during the install. The tested package is at: https://imagemagick.org/archive//binaries/ImageMagick-6.9.13-14-Q16-x64-dll.exe
6) Open up the Windows search and type "path". Click on "Edit the system environment variables".
7) (If you didn't add ImageMagick to the PATH during the install) In the bottom window labelled "System Variables", click Path -> Edit -> Browse -> Look for `C:\Program Files\ImageMagick-X.X.X-QXX` -> Ok
8) In the bottom window labelled "System Variables", click New. Set Variable name to MAGICK_HOME and variable value to `C:\Program Files\ImageMagick-X.X.X-QXX`  
9) In the bottom window labelled "System Variables", click New. Set Variable name to MAGICK_CODER_MODULE_PATH and variable value to `C:\Program Files\ImageMagick-X.X.X-QXX6\modules\coders`
10) Navigate to your PyOCR file install. This will mostly likely be at: `~\AppData\Local\Programs\Python\Python36\Lib\site-packages\pyocr` for Windows. 
11) (Alternatively, if the top didnt work) Open Python and type "import pyocr; print(pyocr.__file__)" and press enter. The file location will print for you.
12) Navigate to the location you found in step 12 on the File Explorer. Edit `builders.py` and `tesseract.py` and replace every instance of "-psm" with "--psm". There should be around three of them.
13) You might need to install GhostScript as well. For the verified version, you can download it using the `gs921w64.exe` file at https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/tag/gs921
14) Add both `C:\Program Files\gs\gs9.22\bin` and `C:\Program Files\gs\gs9.22\lib` to the PATH, in the same way that you added ImageMagick to the path (using the Ghostscript files).
15) Save the file and exit. You're all set to run the invoice sorter now!



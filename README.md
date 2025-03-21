# About PhotoCaptionTool
This script takes a folder full of JPEGs or HEICs and creates a photo log from which you can generate annotated photos or update the originals’ metadata, leveraging [exiftool](https://exiftool.org) to do so. It is written for archaeologist and geared to their needs, but can help with any kind of work that is grouped into projects and sites.

# Installation
## Windows
1. Install [python](https://www.python.org/downloads/).
2. Install [git](https://git-scm.com/downloads).
3. Open a command prompt and run the following command to install PhotoCaptionTool:  
   ```git clone https://github.com/Lugal-PCZ/PhotoCaptionTool.git```
4. cd into the PhotoCaptionTool folder:  
   ```cd PhotoCaptionTool```
5. Install the python dependencies:  
   ```python -m pip install -r requirements.txt```
6. Download and un-zip the [exiftool Windows Executable](https://exiftool.org).
7. Rename the executable from “exiftool -k.exe” to “exiftool.exe”
8. Move exiftool.exe and the “exiftool_files” folder into the PhotoCaptionTool folder.

## MacOS
1. Install [python](https://www.python.org/downloads/).
2. Install [git](https://git-scm.com/downloads).
3. Open a command prompt and run the following command to install PhotoCaptionTool:  
   ```git clone https://github.com/Lugal-PCZ/PhotoCaptionTool.git```
4. cd into the PhotoCaptionTool folder:  
   ```cd PhotoCaptionTool```
5. Install the python dependencies:  
   ```pip3 install -r requirements.txt```
6. Download the [exiftool MacOS Package](https://exiftool.org) and run the installer.

## Linux
PhotoCaptionTool hasn’t yet been tested on Linux, but you should be able to install python, git, and exiftool with your distribution’s package manager. Then follow steps 3–5 of the MacOS instructions, above.

# Updating PhotoCaptionTool
1. Open a command prompt and cd into the PhotoCaptionTool folder.
2. Run ```git pull```.

# Usage
## Basic Instructions
1. Open a command prompt and cd into the PhotoCaptionTool folder.
2. Run PhotoCaptionTool with the following command:  
   (_Windows_) ```python photo_caption_tool.py```  
   (_MacOS / Linux_) ```python3 photo_caption_tool.py```
   You will be presented with PhotoCaptionTool’s main menu.
3. (_optional_) Enter E to edit PhotoCaptionTool’s presets (_see below_).
4. Enter 1 to load a folder full of JPEGs.
5. Type the path to the folder of JPEGs or drag that folder onto the command window, then press enter.
6. Enter 2 to create a photo log.
7. Enter 3 or double-click on the “Photo Log.csv” file in the source folder to view and edit the photo log.
8. Edit the photo log as necessary and save your changes.
9. (_optional_) Enter 4 to copy the the unmodified JPEGs into a “Renamed Photos” folder, with the naming convention indicated in PhotoCaptionTool’s presets (_see below_). HEIC files will be automatically converted to JPEGs.
10. (_optional_) Enter 5 to created annotated versions of the JPEGs, saved to an “Annotated Photos” folder, with the naming convention indicated in PhotoCaptionTool’s presets (_see below_). HEIC files will be automatically converted to JPEGs.
11. (_optional_) Enter 6 to create a Word doc with the photos with captions taken from the photo log. You can then enter 7 or double-click the Word doc in the source folder to view it.
12. Repeat steps 3 through 11 as necessary, or enter Q to quit.

## Editing PhotoCaptionTool’s Presets
The first time that you run PhotoCaptionTool it will create a generic configs.ini file in the PhotoCaptionTool folder. You can then edit that configs.ini file directly prior to running PhotoCaptionTool or press E at its main menu to preset data and tailor the way that the script generates its outputs. The available options are:
* **exiftool**  
  Edit this value to change the path to exiftool, if it is not installed in the default location.

* **papersize**  
  Edit this value to change the paper size of the Word doc from A4 to Letter.

* **subjectdelimiter**  
  PhotoCaptionTool will extract a subject from the photos’ captions (when, for example, they’re taken on an iPad). By default, the subject precedes the first colon in the caption, but this value can be changed to a delimiter of your choosing. If no subject is found in the original JPEGs, the “Subject” column will be left blank in the photo log.

* **photographer**  
  Edit this value to pre-set the name or initials of the photographer in the photo log.

* **project**  
  Edit this value to pre-set the name of the project in the photo log.

* **site**  
  Edit this value to pre-set the name of the site in the photo log.

* **precision**  
  Edit this value to change the way that the directions that the photographs are facing (as captured by the device’s GPS) are reported in the photo log. The options are:  
    * _coarse_ (N, NE, E, SE, S, SW, W, NW)
    * _fine_ (N, NNE, NE, ENE, E, and so-on)
    * _precise_ (the actual bearing, in degrees)

* **format**  
  Edit this value to change the way that files are renamed. The options are:
    * _1_ (Subject -- Photographer_Photo.jpg)
    * _2_ (Site_Subject_Sequence.jpg)

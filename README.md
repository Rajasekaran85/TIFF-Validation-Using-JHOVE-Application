# Application Title

TIFF Validation Using JHOVE Application

## Description

* Validating the TIFF images using the JHOVE application
* **JHOVE:** Open source file format identification, validation & characterisation
* Output information will be written in the excel file
* Input files should be TIFF file path
* Output log file will written in the same path of the TIFF file with the below fields
* **Properties:** Status (file is valid or not), TIFF Version, Compression, Width & Height, Color Mode, ICC Profile, Date Time, Artist, Scanner Name, Scanner Model, Scanner Software, Orientation, Resolution, Bits Per Sample, Sample Per Pixel 

## Getting Started

### Dependencies

* Windows 7 or above

### Installing

* xlsxwriter library installed
* Install JHOVE application from the https://jhove.openpreservation.org/

### Executing program

* Capture the JHOVE application installed location path in the INI file, e.g.: <path>D:\JHOVE\jhove.bat</path>
* Run the program
* Tool will ask to enter the Input file path of the input TIFF file path
* Tool executes the TIFF files and created the "Validation_log.xlsx" file. 

## Version History

* 0.1
    * Initial Release

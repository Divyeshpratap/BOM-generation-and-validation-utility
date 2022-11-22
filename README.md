# BOM-generation-and-validation-utility
This VBA utility aims at Generating End User understandable Sub Assembly BOM from the PSV value (refer annexures for more details on what is PSV file type, and how is it generated) provided by the user.
Utility Functionality
Input Variable -> AutoBOM generated PSV file of Main Assembly, TcM id of assemblies for which sub assembly BOM is required by User
Output -> BOM of Main assembly, Separate BOM of individual assemblies


Steps for generating Sub Assembly BOM are:
1. User inputs the file location of PSV file in sheet: , cell number:
2. User presses command button ( Generate BOM).
3. Utility goes into processing mode and successfull BOM generation is informed to user through a msgbox "BOM generation complete".


Note: VBA Macros should be enabled in excel instance. Minimum 16GB RAM required. MS office version 2010 or higher required. All other excel files should be closed before running the utility. Multiple instances of excel shoould not be opened.


To understand the functionality and troubleshooting the code, it is necessary to get a deep understanding of the input (PSV file) and internal assembly structureof a Solid Edge Assembly file.  This section aims at providing a brief overview of the same. A Solid Edge assembly is generated at multiple levels where the Main assembly is given the lowest level of 1 and subsequent parts\ assemblies are assigned levels in ascending order. These levels are defined by software to maintain the parent\ child relationship of various components used in Main assembly.
These levels are automatically assigned by SE, and a thorough understanding of assembly structure can be realized using TcM structure manager.

To generate BOM in excel format which is a part of ATE workflow; it is necessary that the part list of components is fetched from within Solid Edge and is stored in an intermediate file.
Here, for this particular purpose a utility developed by Siemens is used which can retreive part list from SE environment using SE API's embedded in visual basic.
This utility retreives certain parameters (these parameters can be referred from psv file first row) of all parts used in main assembly in a sequential manner as per their sequential arrangement in structure manager.
After running the Siemens AutoBOM utility, the PSV file is by default generated and stored at "D:\BOM_Report_Files location" in respective PC. PSV is a pipe separated file type which is similar to ".csv" file with the only difference being the deliminator in ".psv" file type is a pipe "|" character.

This utility makes use of the same PSV file and converts it into a user readable and manipulative excel document for further reference and communication.

Steps internally executed by PSV file are:
1. Utility imports psv file into excel into a separate sheet "Psv_values".
2. Utility reads the list of sub assembly id provided by user, checks for correctness of assembly id and then generates seperate "Sub_psv" sheet for the first sub assembly.
3. For the first sub assembly, code starts segregating the BOM into separate sheets ( "Maintenance", "BOP", "Machined", "Inhouse")  based on their ALB component type.
4. The values are populated and all the four sheets are tranferred to a new workbook, which gets placed in the sub directory "Sub" and get renamed as per the TcM id.
5. The process is repeated for all sub assemblies for which separate BOM was requested by the user.
6. After all sub assembly BOM's are generatd utility moves to generate the BOM of master assembly. The master assembly BOM is the main assembly BOM minus the sum of individual sub assembly BOM.
7. This is performed by eliminating the sub assembly values from original "Psv_values" sheet. Master asembly BOM is generated and exported within the same directory.
8. After completion of all above sub routines, notification for process complete is shown.
  
For any iterations in utility contact: ATE 

-Divyesh
28th July 2021

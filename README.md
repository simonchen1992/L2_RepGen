# RRAP Checktool_Vendor #
RRAP checktool is used to check Device ICS configuration which is provided by VISA. This tool mainly provides two main functions.
## Generate RRAP from ICS automatically ##
You can generate RRAP from origin ICS by following instruction.

1. Decrypt ICS by qpdf with following command, qpdf --password= --decrypt PathofInputFile PathofOutputFile
2. Execute chektool.exe
3. Following prompt instruction and input the path of decrypted ICS
4. RRAP is generated
## Compare the result of two RRAP ##
Sometimes it's good to have a double check with manually filled RRAP and automatically filled RRAP.

1. Choose the function of comparing two RRAP
2. Remove the data validation format of manually filled RRAP.
3. enter the path of both RRAP.
4. Comparing result will be shown in the console.
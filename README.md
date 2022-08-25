# Convert-OfficeFiles

List of functions all aimed to convert old Office files, into the newer format.
ConvertDOC-ToDOCX

ConvertXLS-ToXLSX

ConvertPPT-ToPPTX

Convert-OfficeFiles

  - This is a reference function for the rest of the functions - all in one.
  
  
-------------------------------------------------------------------------------------------------
Each function provides a -Path, -Destination, -RemoveOld, and -Recurse parameter. 
  - The -Path parameter takes a literal path of a file and/or, directory.
    - Specifying a file path will convert the file by itself.
    - Specifying a directory will convert all files found in the directory.
    
  - Using -Destination allows you to specify where the converted files should be placed.
  - When -RemoveOld is used, the old file will be deleted when the conversion has occurred.
  - If a file path is specified, the -Recurse parameter will not show list as an option to use.

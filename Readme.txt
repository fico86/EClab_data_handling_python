Sorting CV data from EC-lab by cycles and normalised by mass, to plot in Origin.

Input format: 
Excel file containing 3 sheets named "cv_files", "mass_grid", "selected_cycles"
eg. TiN + Carbon ink mixture vs carbon control ink mixture, both cycled in Ar and then O2

"cv_files" sheet:
    columns: list of treatment on samples, eg. cycled in Ar, then O2
    rows: list of diffrent samples, eg. TiN + carbon, carbon control
    values:cv file names of with extention of the cooresponding CVs
-------------------------------------------------------------------------------
        | Ar                               | O2                               |
-------------------------------------------------------------------------------
TiN + C | (Ar cv file name with extention) | (Ar cv file name with extention) |
-------------------------------------------------------------------------------
C       | (Ar cv file name with extention) | (Ar cv file name with extention) |
-------------------------------------------------------------------------------

"mass_grid" sheet:
    columns: list diffrent masses to be normalised by 
    rows: list of diffrent samples, must match the rows in "cv_files"
    values: corresponding masses in grams, masses set to 0 will not be sorted
---------------------------------------------------------------------
        | active mass            | Tin           | carbon           |
---------------------------------------------------------------------
TiN + C | (mass of TiN + Carbon) | (mass of TiN) | (mass of carbon) |
---------------------------------------------------------------------
C       | (mass of Carbon)       | 0             | (mass of Carbon) |
---------------------------------------------------------------------

"selected_cycles'
    columns: list of treatment on samples, must match columns of "cv_files"
    rows: general list of numbers starting from 0, does not effect the sorting
    values: cycle numbers to be sorted, for each treatment, set to 0 finishes cycle selection
--------------
  | Ar | O2  |
--------------
0 | 1  | 1   |
--------------
1 | 3  | 5   |
--------------
2 | 0  | 10  |
--------------
3 | 0  | 15  |
--------------
4 | 0  | 20  |
--------------

In output file, will be sorted into sheets based on the normalising masses. 
The first sheet is 'as measured', ie. not normalised, followed by the masses listed in the "mass_grid". 
When mass is set to 0, that particuler normalisation is ignored, ie the TiN normalised sheet, will only have the TiN + C CV values.

===============================================================================================================================================================

Sorting galvanostatic data from EC-lab by cycles and normalised by mass, to plot in Origin.

Input format: 
Excel file containing 3 sheets named "gal_files", "mass_grid", "selected_cycles"
eg. TiN + Carbon ink mixture vs carbon control ink mixture, both cycled in Ar and then O2

"gal_files" sheet:
    columns: list of treatment on samples, eg. cycled in Ar, then O2
    rows: list of diffrent samples, eg. TiN + carbon, carbon control
    values:cv file names of with extention of the cooresponding CVs
-------------------------------------------------------------------------------
        | Ar                               | O2                               |
-------------------------------------------------------------------------------
TiN + C | (Ar cv file name with extention) | (Ar cv file name with extention) |
-------------------------------------------------------------------------------
C       | (Ar cv file name with extention) | (Ar cv file name with extention) |
-------------------------------------------------------------------------------

"mass_grid" sheet:
    columns: list diffrent masses to be normalised by 
    rows: list of diffrent samples, must match the rows in "gal_files"
    values: corresponding masses in grams, masses set to 0 will not be sorted
---------------------------------------------------------------------
        | active mass            | Tin           | carbon           |
---------------------------------------------------------------------
TiN + C | (mass of TiN + Carbon) | (mass of TiN) | (mass of carbon) |
---------------------------------------------------------------------
C       | (mass of Carbon)       | 0             | (mass of Carbon) |
---------------------------------------------------------------------

"selected_cycles'
    columns: list of treatment on samples, must match columns of "gal_files"
    rows: general list of numbers starting from 0, does not effect the sorting
    values: cycle numbers to be sorted, for each treatment, set to 0 finishes cycle selection
--------------
  | Ar | O2  |
--------------
0 | 1  | 1   |
--------------
1 | 3  | 5   |
--------------
2 | 0  | 10  |
--------------
3 | 0  | 15  |
--------------
4 | 0  | 20  |
--------------

In output file, will be sorted into sheets based on the normalising masses. 
The first file is 'as measured', ie. not normalised, followed by the masses listed in the "mass_grid". 
When mass is set to 0, that particuler normalisation is ignored, ie the TiN normalised sheet, will only have the TiN + C gal values.
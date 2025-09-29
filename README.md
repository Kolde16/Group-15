# BIManalyst group 15

## Focus area: Indoor

The claim we are checking: We are checking the U-value information of the buidlings encvelope, this includes the windows, walls, ceilings and floors. Doors and other building elements are not included in our analysis, as they are not present in the MEP report.

The first assignment will only include the windows!

The claims are present in the 25-16-D-MEP report on page 5 and 6 in section 2,1 and 2,4, specifically in table 1 and 3. 

## Script discription

This script reads windows from an **IFC model** and exports their main properties to Excel.

### What it does
- Gets window **dimensions** (width, height, area).
- Reads **frame width** and **rebate depth** (if available).
- Finds the **glass** and **frame materials** and their thermal conductivity.
- Calculates:
  - **Glass area** and **frame area**
  - **Glass U-value** (from conductivity)
  - **Frame U-value** (conductivity ÷ thickness, default k=0.2 if missing)
  - **Overall window U-value** as an area-weighted average

### Output
An Excel file is created on your Desktop with one row per window, including:

- GlobalId and Name  
- Width, Height, Total Area  
- Frame width, rebate depth, and frame area  
- Glass and frame materials with conductivity  
- U-values for glass, frame, and the whole window  

TEST
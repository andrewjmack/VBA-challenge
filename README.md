#  VBA Challenge
#   Andrew Mack | March 2024

## Contents of Repo
  VBA script file
  Screenshots of worksheets resulting from script
  README.md

## Resources
  Original .xlsm files and instructions provided by the University.
  Referenced prior classroom activities and lesson recordings to help with VBA scripting to meet challenge requirements
  Microsoft online support documentation for worksheet looping: https://support.microsoft.com/en-us/help/142126/macro-to-loop-through-all-worksheets-in-a-workbook
  Online source that helped with scripting min and max functions: https://www.wallstreetmojo.com/vba-max/
    
## Challenges & Observations
  Wary of date column formatting in data set: some cells stored as text rather than numbers...
  Encountered div/0 and runtime 6 errors calculating the annual price % change...
    initially avoided by nesting if statement for when annual opening price = 0, to return a comment rather than attempt to calculate...
    ...realized something in my script had beeb reseting the initial annual opening price
    UPDATE: fixed by reversing "Cells(2, 3).Value = Open_Annual" to "Open_Annual = Cells(2, 3).Value"
  Struggled with For Each loop across sheets, until realizing that all cell and range objects needed to start with "ws."

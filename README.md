# Excel Form View

![](https://raw.githubusercontent.com/kking423/excel-form-view/main/readme-resources/Form-View.png)


## The Problem:
Workbooks can often get cluttered with multiple sheets containing huge amounts of data with dozens or more columns. 
It can be brutal trying to scroll down, up, and across the sheet specially when contending with 
different row, column, and cell formats. Excel does provide a built-in feature with a pop-up form, 
but it's not super intuitive to use, it's bound to only 32 fields, doesn't carry-over validation 
lists and formulas, and doesn't work with the Mac Version.  

![](https://raw.githubusercontent.com/kking423/excel-form-view/main/readme-resources/Data-Table-Challenges.png)

## The Solution:
Using the built-in Form View as a guide, 
I developed a similar feature using VBA that dynamically generates a vertical/transposed view of any row from a 
data table object.  Simply double-click within a cell of the row you want to view, and the 
Form-View sheet immediately loads. From there you can make in-line updates and save back to the original source table. 

I've often use this custom Form-View as a quick way to manage small datasets, prototypes, 
and a tool to make it easier to present data during screen shares. This has saved me tremendous amounts of time 
over the years and has received a lot of great feedback from co-workers. It's not a feature I use all the time, 
but when I need it, I find that it's an invaluable utility that makes it easier to work with raw data.

* When enabled, you can double-click any cell in any Data Table in Excel workbook and launch the Form-View sheet.
* Within Form-View, you can make in-line updates back to the source sheet
* All the source table formulas, validations, conditional formatting, notes, etc. is preserved in Form-View

## Setup:
* Step 1. Download and open the Form-View-NonMacro-Starter.xlsx workbook. Save a new copy as a macro-enabled version. None of the code will work until we import the module in the next step.
* Step 2. Download the mod_FormView.bas file. Within the "Developer" menu, import module into your newly created macro-enabled workbook.
* Step 3. After that, you'll also need to add a single line of code (shown below) to your ***Workbook_SheetBeforeDoubleClick*** method (Workbook module).
    
    ```
        Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
            Call FormView_Main
        End Sub
    ```
  
  * Note: For Mac versions, you may have to copy-paste the entire procedure shown below if you get an error trying to access the module.

* Step 4. Go to the Form-View sheet. You will need to assign a macro to each of the form-controls (there are 5 of them shown below). The names of the macros below should appear in the list if you imported the module file in Step 2.
    
    ```
      * Checkbox: Inherit Source Data Formatting --> FormView_Toggle_Formatting
      * Checkbox: Enable Cell Double Click --> FormView_Toggle_DblClick
      * Button: Highlight Special Cells --> FormView_Highlight_SpecialCells
      * Button: Highlight Blanks --> FormView_Highlight_Blanks
      * Button: Save Changes --> FormView_Save
    ```

* Step 5. The starter workbook includes a few different sample data sets to illustrate how the Form-View works with different data tables.
    * Note: The Form-View will only work when data is assigned to a named Data Table object.
      

  

      
  


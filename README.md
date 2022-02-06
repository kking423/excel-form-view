# Excel Form View
Excel already has a built-in "Form View" feature, but it does have some limitations for wide tables (lots of columns) 
and it doesn't work on Mac versions. With a little VBA magic, we can create our own version of "Form View" that makes
reviewing row-level data much more enjoyable and practical.

## What it Does:
* When enabled, you can double-click any cell in any Data Table in Excel workbook and launch the Form-View sheet.
* Within Form-View, you can make in-line updates back to the source sheet
* All the source table formulas, validations, conditional formatting, notes, etc. is preserved in Form-View

## Setup:
* Step 1. Download and open the Form-View-NonMacro-Starter.xlsx workbook. Save a new copy as a macro-enabled version. None of the code will work until we import the module in the next step.\
* Step 2. Download the mod_FormView.bas file. Within the "Developer" menu, import module into your newly created macro-enabled workbook.\
* Step 3. After that, you'll also need to add a single line of code (shown below) to your ***Workbook_SheetBeforeDoubleClick*** method (Workbook module).\ 
    
    ```
        Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
            Call FormView_Main
        End Sub
    ```
  
  * Note: For Mac versions, you may have to copy-paste the entire procedure shown below if you get an error trying to access the module.\
\
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
      

  

      
  


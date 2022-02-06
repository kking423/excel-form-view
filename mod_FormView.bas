Attribute VB_Name = "mod_FormView"
Option Explicit

Dim src_sheet As String
Dim tbl_name As String
Dim tbl As ListObject
Dim header_addr As String
Dim row_addr As String
Dim selected_row As Long
Dim sub_addr1 As String
Dim sub_addr2 As String
Dim sub_addr3 As String
Dim col_count As Integer
Dim first_cell As String
Dim form_data_full_range As String
Dim form_data_value_range As String
Dim form_data_field_range As String
Dim form_data_start_cell As String
Dim form_data_start_row As Integer
Dim form_value_start As String
Dim form_data_value_column As String
Dim form_data_last_row As Integer
Dim form_data_row_range As String
Dim src_first_col As String


Sub FormView_Main()
Dim config() As String

    config = FormView_Config_Setting("Form-Config", "CLICK")
    
    If Not Selection.ListObject Is Nothing And config(1) = True And ActiveSheet.name <> "Form-View" Then
        Call FormView_Display
    End If
    
End Sub


Sub FormView_Display()
Dim config() As String

On Error GoTo Err
    Call FormView_Calculate_Source_Settings

    Application.ScreenUpdating = False
        Call FormView_Load

        config = FormView_Config_Setting("Form-Config", "FORMATTING")

        If config(1) = False Then
            Call FormView_Formats 'optional

        End If

        'set the active cell to make it easier for user to make any updates (if needed)
        Sheets("Form-View").Range(form_value_start).Select

    Application.ScreenUpdating = True

Err:
    If Err.Number <> 0 Then
        MsgBox "[" & Err.Number & "] " & Err.Description & vbCrLf & vbCrLf & "Source: " & Err.Source, vbExclamation, "Data Form View: Error"
        Err.Clear
        Resume Next
    End If
End Sub





'-------------------------------------------------------------------------------------------
'User Customizations
'-------------------------------------------------------------------------------------------
Sub FormView_Toggle_Formatting()
Dim config() As String

    config = FormView_Config_Setting("Form-Config", "FORMATTING")
    
    If config(1) = True Then
        MsgBox "Unchecking will attempt to optimize the formatting in Form-View but may load more slowly." & vbCrLf & vbCrLf & _
            "This setting will take effect the next time you load the Form-View.", vbInformation, "Form-View | Optimized Formatting (Slower)"
    Else
        MsgBox "This is the faster (default) option but reduce readability in the Form-View." & vbCrLf & vbCrLf & _
            "This setting will take effect the next time you load the Form-View.", vbInformation, "Form-View | Default Formatting (Faster)"
    End If
    
    Sheets("Form-Config").Range(config(0)).Value = FormView_Toggle_Boolean(config(1))
    
End Sub

Sub FormView_Toggle_DblClick()
Dim config() As String

    config = FormView_Config_Setting("Form-Config", "CLICK")
    Sheets("Form-Config").Range(config(0)).Value = FormView_Toggle_Boolean(config(1))
    
End Sub




'-------------------------------------------------------------------------------------------
'Calculations - Used to Retrieve and Place Data on the Form
'-------------------------------------------------------------------------------------------
Sub FormView_Calculate_Source_Settings()
    src_sheet = ActiveSheet.name
    tbl_name = Selection.ListObject.name
    Set tbl = ActiveSheet.ListObjects(tbl_name)
    selected_row = ActiveCell.Row
    header_addr = tbl.HeaderRowRange.Address(False, False)
    row_addr = Replace(header_addr, tbl.HeaderRowRange.Row, selected_row, , , vbTextCompare)
    first_cell = Split(row_addr, ":")(0)
    src_first_col = FormView_ColumnLetter(Range(first_cell).Column)
    sub_addr1 = "'" & src_sheet & "'!A1"
    sub_addr2 = "'" & src_sheet & "'!" & src_first_col & selected_row
    sub_addr3 = "'" & src_sheet & "'!" & ActiveCell.Address
    col_count = Range(tbl.HeaderRowRange.Address).Columns.Count
    
    form_data_start_cell = FormView_Config_Setting("Form-Config", "START_CELL")(1)
    form_data_start_row = Range(form_data_start_cell).Row
    form_value_start = Replace(Range(form_data_start_cell).Offset(0, 1).Address, "$", "", compare:=vbTextCompare)
    form_data_value_column = Left(form_value_start, 1)
    form_data_last_row = col_count + form_data_start_row - 1
    
    form_data_full_range = form_data_start_cell & ":" & form_data_value_column & form_data_last_row
    form_data_field_range = form_data_start_cell & ":" & Left(form_data_start_cell, 1) & form_data_last_row
    form_data_value_range = form_value_start & ":" & Left(form_value_start, 1) & form_data_last_row
    form_data_row_range = CStr(form_data_start_row) & ":" & CStr(form_data_last_row)
End Sub



'-------------------------------------------------------------------------------------------
'Load Data
'-------------------------------------------------------------------------------------------

Sub FormView_Load()
Dim rng As Range

    '------Load Source Data
    Sheets("Form-View").Select
    Sheets("Form-View").Cells.Clear
    
    Set rng = Sheets("Form-View").Range(form_data_start_cell)
    
    Sheets(src_sheet).Range(header_addr & "," & row_addr).Copy
    rng.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False
    
    '------Add Header
    Set rng = Range(form_data_start_cell).Offset(-5, 0)
    With rng
        .Value = "Form-View"
        .Font.Size = 24
        .Font.name = "Arial"
        .Font.Bold = True
    End With
    
    '------Add Navigation back to source sheet
    Set rng = Range(form_data_start_cell).Offset(-4, 0)
    With rng
        rng.Hyperlinks.Add Anchor:=rng, Address:="", SubAddress:=sub_addr1, TextToDisplay:="Back to " & src_sheet
    End With
    
    Set rng = Range(form_data_start_cell).Offset(-3, 0)
    With rng
        rng.Hyperlinks.Add Anchor:=rng, Address:="", SubAddress:=sub_addr2, TextToDisplay:="Back to Selection (Start of Row# " & selected_row & ")"
    End With
    
    Set rng = Range(form_data_start_cell).Offset(-2, 0)
    With rng
        rng.Hyperlinks.Add Anchor:=rng, Address:="", SubAddress:=sub_addr3, TextToDisplay:="Back to Selection (Row# " & selected_row & ")"
    End With
    
    
End Sub



'-------------------------------------------------------------------------------------------
'Save
'-------------------------------------------------------------------------------------------

Sub FormView_Save()
Dim dest As Range
Dim src As Range

    If src_sheet <> "" Then
        Set src = Sheets("Form-View").Range(form_data_value_range)
        src.Copy
        Set dest = Sheets(src_sheet).Range(first_cell)
        src.Copy 'redundant step seemed to be needed to work on Mac version
        
        With dest
            .PasteSpecial XlPasteType.xlPasteValuesAndNumberFormats, xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=True
            .PasteSpecial XlPasteType.xlPasteComments, xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=True
            .PasteSpecial XlPasteType.xlPasteValidation, xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=True
            .PasteSpecial XlPasteType.xlPasteFormulasAndNumberFormats, xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=True
        End With
        
        Application.CutCopyMode = False
        MsgBox "Updates were applied to the source sheet", vbInformation, Title:="Form View Changes"
        
    End If
    
    Application.ScreenUpdating = True
    
End Sub


'-------------------------------------------------------------------------------------------
'Formatting
'-------------------------------------------------------------------------------------------

Sub FormView_Formats()
Dim rng As Range


    '-----Format Rows
    Set rng = Range(form_data_row_range)
    With rng
        .RowHeight = 30
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    '-------Format Field Names
    Set rng = Range(form_data_field_range)
    With rng
        .Interior.Pattern = xlSolid
        .Interior.PatternThemeColor = xlThemeColorAccent1
        .Interior.ThemeColor = xlThemeColorDark1
        .Interior.TintAndShade = -0.0499893185216834
        .Interior.PatternTintAndShade = 0
        
        .Font.ThemeColor = xlThemeColorLight1
        .Font.TintAndShade = 0
    End With
    
    '-------Format Form Data Values
    Set rng = Range(form_data_value_range)
    With rng
        .Interior.Pattern = xlNone
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        
        .Font.ThemeColor = xlThemeColorLight1
        .Font.TintAndShade = 0
    End With
    
End Sub


Sub FormView_Highlight_Special_Cells()
On Error GoTo Err 'necessary to skip over warnings when no items are found

Dim rng As Range
    Set rng = Range(form_data_value_range)
    
    With rng.SpecialCells(xlCellTypeFormulas).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 14414040
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With rng.SpecialCells(xlCellTypeComments).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 14414040
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With rng.SpecialCells(xlCellTypeAllValidation).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 14414040
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

Err:
    Err.Clear
    Resume Next
End Sub

Sub FormView_Highlight_Blanks()
On Error GoTo Err 'necessary to skip over warnings when no items are found

Dim rng As Range
    Set rng = Range(form_data_value_range)
    
    With rng.SpecialCells(xlCellTypeBlanks).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 14548991
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
Err:
    Err.Clear
    Resume Next
End Sub



'-------------------------------------------------------------------------------------------
'Functions
'-------------------------------------------------------------------------------------------
Function FormView_Config_Setting(src_sheet As String, keyword As String)
Dim config(0 To 1) As String
Dim rng As Range

    Set rng = Sheets(src_sheet).Columns("A").Find(what:=keyword, LookIn:=xlValues, lookat:=xlPart)
       
    If Not rng Is Nothing Then
        config(0) = Sheets(src_sheet).Range(rng.Address).Offset(0, 1).Address
        config(1) = Sheets(src_sheet).Range(rng.Address).Offset(0, 1).Value
    End If
    
    FormView_Config_Setting = config
End Function

Function FormView_Toggle_Boolean(val) As Boolean
'used to reverse a form control setting
    If val = False Then
        val = True
    Else
        val = False
    End If
    
    FormView_Toggle_Boolean = val
End Function

Function FormView_ColumnLetter(ByVal col As Long) As String
Dim i1 As Long
Dim i2 As Long

    i1 = (col - 1) \ 26   ' col - 1 =i1*26+i2 : this calculates i1 and i2 from col
    i2 = (col - 1) Mod 26
    FormView_ColumnLetter = Chr(Asc("A") + i2) ' if i1 is 0, this is the column from "A" to "Z"
    If i1 > 0 Then 'in this case, i1 represents the first letter of the two-letter columns
        FormView_ColumnLetter = Chr(Asc("A") + i1 - 1) & FormView_ColumnLetter ' add the first letter to the result
    End If
End Function





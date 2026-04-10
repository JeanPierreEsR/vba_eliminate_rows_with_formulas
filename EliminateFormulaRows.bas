Attribute VB_Name = "EliminateFormulaRows"
' ============================================================
'  EliminateFormulaRows.bas
'  Deletes every row that contains at least one formula cell
'  within the currently selected range.
'
'  Usage:
'    1. Select the cells you want to inspect on the worksheet.
'    2. Open the VBA editor (Alt + F11).
'    3. Insert > Module and paste this code, or import this .bas file.
'    4. Run DeleteRowsWithFormulas from the Macros dialog (Alt + F8).
' ============================================================

Option Explicit

' ------------------------------------------------------------
'  DeleteRowsWithFormulas
'  Scans the current selection and deletes any row that
'  contains one or more formula cells within that selection.
'
'  Parameters:
'    selRng (optional) – range to inspect; defaults to Selection.
' ------------------------------------------------------------
Public Sub DeleteRowsWithFormulas(Optional selRng As Range = Nothing)

    Dim scanRng      As Range
    Dim formulaCells As Range
    Dim formulaCell  As Range
    Dim rowsToDelete As Range

    ' --- resolve the range to scan --------------------------------------
    If selRng Is Nothing Then
        If TypeName(Selection) <> "Range" Then
            MsgBox "Please select a range of cells before running this macro.", _
                   vbExclamation
            Exit Sub
        End If
        Set scanRng = Selection
    Else
        Set scanRng = selRng
    End If

    ' --- find all formula cells within the selection --------------------
    On Error Resume Next
    Set formulaCells = scanRng.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If formulaCells Is Nothing Then
        MsgBox "No formula cells found in the selection. No rows were deleted.", _
               vbInformation
        Exit Sub
    End If

    ' --- collect the unique rows that hold at least one formula cell ----
    '     Union the entire rows so all deletions happen in one call,
    '     avoiding row-index shifting issues.
    For Each formulaCell In formulaCells
        If rowsToDelete Is Nothing Then
            Set rowsToDelete = formulaCell.EntireRow
        Else
            Set rowsToDelete = Union(rowsToDelete, formulaCell.EntireRow)
        End If
    Next formulaCell

    ' --- count distinct rows --------------------------------------------
    Dim rowCount As Long
    Dim area     As Range
    For Each area In rowsToDelete.Areas
        rowCount = rowCount + area.Rows.Count
    Next area

    ' --- confirm before deleting ----------------------------------------
    Dim answer As VbMsgBoxResult
    answer = MsgBox( _
        "This will permanently delete " & rowCount & " row(s) on sheet """ & _
        scanRng.Worksheet.Name & """." & vbCrLf & vbCrLf & "Continue?", _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "Delete rows with formulas")

    If answer = vbNo Then
        MsgBox "Operation cancelled. No rows were deleted.", vbInformation
        Exit Sub
    End If

    ' --- delete the rows ------------------------------------------------
    rowsToDelete.Delete Shift:=xlShiftUp

    MsgBox rowCount & " row(s) deleted successfully.", vbInformation

End Sub


' ------------------------------------------------------------
'  DeleteRowsWithFormulas_NoPrompt
'  Silent version – no confirmation dialog.  Useful when
'  called from another macro or when you want no UI at all.
'
'  Parameters:
'    selRng (optional) – range to inspect; defaults to Selection.
'
'  Returns the number of rows that were deleted.
' ------------------------------------------------------------
Public Function DeleteRowsWithFormulas_NoPrompt( _
        Optional selRng As Range = Nothing) As Long

    Dim scanRng      As Range
    Dim formulaCells As Range
    Dim formulaCell  As Range
    Dim rowsToDelete As Range
    Dim rowCount     As Long

    If selRng Is Nothing Then
        If TypeName(Selection) <> "Range" Then
            DeleteRowsWithFormulas_NoPrompt = 0
            Exit Function
        End If
        Set scanRng = Selection
    Else
        Set scanRng = selRng
    End If

    On Error Resume Next
    Set formulaCells = scanRng.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If formulaCells Is Nothing Then
        DeleteRowsWithFormulas_NoPrompt = 0
        Exit Function
    End If

    For Each formulaCell In formulaCells
        If rowsToDelete Is Nothing Then
            Set rowsToDelete = formulaCell.EntireRow
        Else
            Set rowsToDelete = Union(rowsToDelete, formulaCell.EntireRow)
        End If
    Next formulaCell

    Dim area As Range
    For Each area In rowsToDelete.Areas
        rowCount = rowCount + area.Rows.Count
    Next area

    rowsToDelete.Delete Shift:=xlShiftUp

    DeleteRowsWithFormulas_NoPrompt = rowCount

End Function

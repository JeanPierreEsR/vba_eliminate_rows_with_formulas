Attribute VB_Name = "EliminateFormulaRows"
' ============================================================
'  EliminateFormulaRows.bas
'  Deletes every row in the used range that contains at least
'  one cell with a formula.
'
'  Usage:
'    1. Open the VBA editor (Alt + F11).
'    2. Insert > Module and paste this code, or import this .bas file.
'    3. Run DeleteRowsWithFormulas from the Macros dialog (Alt + F8).
' ============================================================

Option Explicit

' ------------------------------------------------------------
'  DeleteRowsWithFormulas
'  Scans the active sheet's used range and deletes any row
'  that contains one or more formula cells.
'
'  Parameters:
'    ws  (optional) – target Worksheet; defaults to ActiveSheet.
' ------------------------------------------------------------
Public Sub DeleteRowsWithFormulas(Optional ws As Worksheet = Nothing)

    Dim targetSheet  As Worksheet
    Dim usedRng      As Range
    Dim formulaCells As Range
    Dim formulaCell  As Range
    Dim rowsToDelete As Range

    ' --- resolve target sheet -------------------------------------------
    If ws Is Nothing Then
        Set targetSheet = ActiveSheet
    Else
        Set targetSheet = ws
    End If

    ' --- locate the used range ------------------------------------------
    Set usedRng = targetSheet.UsedRange
    If usedRng Is Nothing Then
        MsgBox "The sheet is empty. Nothing to do.", vbInformation
        Exit Sub
    End If

    ' --- find all cells that contain a formula --------------------------
    On Error Resume Next
    Set formulaCells = usedRng.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If formulaCells Is Nothing Then
        MsgBox "No formula cells found. No rows were deleted.", vbInformation
        Exit Sub
    End If

    ' --- collect the unique rows that hold at least one formula cell ----
    '     Walk the formula cells and union their entire rows so we can
    '     delete them all in one shot (avoids row-index shifting issues).
    For Each formulaCell In formulaCells
        If rowsToDelete Is Nothing Then
            Set rowsToDelete = formulaCell.EntireRow
        Else
            Set rowsToDelete = Union(rowsToDelete, formulaCell.EntireRow)
        End If
    Next formulaCell

    ' --- confirm before deleting ----------------------------------------
    Dim rowCount As Long
    rowCount = 0

    ' Count distinct rows in the union range
    Dim area As Range
    For Each area In rowsToDelete.Areas
        rowCount = rowCount + area.Rows.Count
    Next area

    Dim answer As VbMsgBoxResult
    answer = MsgBox( _
        "This will permanently delete " & rowCount & " row(s) on sheet """ & _
        targetSheet.Name & """." & vbCrLf & vbCrLf & "Continue?", _
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
'  Returns the number of rows that were deleted.
' ------------------------------------------------------------
Public Function DeleteRowsWithFormulas_NoPrompt( _
        Optional ws As Worksheet = Nothing) As Long

    Dim targetSheet  As Worksheet
    Dim usedRng      As Range
    Dim formulaCells As Range
    Dim formulaCell  As Range
    Dim rowsToDelete As Range
    Dim rowCount     As Long

    If ws Is Nothing Then
        Set targetSheet = ActiveSheet
    Else
        Set targetSheet = ws
    End If

    Set usedRng = targetSheet.UsedRange
    If usedRng Is Nothing Then
        DeleteRowsWithFormulas_NoPrompt = 0
        Exit Function
    End If

    On Error Resume Next
    Set formulaCells = usedRng.SpecialCells(xlCellTypeFormulas)
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

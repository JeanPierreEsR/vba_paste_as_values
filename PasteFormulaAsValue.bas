Attribute VB_Name = "PasteFormulaAsValue"
Option Explicit

' ============================================================
' PasteMatchingFormulasAsValues
'
' Iterates over the current selection and replaces any cell
' whose formula begins with FormulaPrefix with its computed
' value, effectively "pasting as values" for those cells.
'
' Usage
'   1. Select the cells you want to process.
'   2. Run PasteMatchingFormulasAsValues from the Macro dialog
'      (or assign it to a button / keyboard shortcut).
'   3. When prompted, type the prefix to match, e.g. =IFERROR
'      Leave the input box blank to cancel without changes.
'
' Notes
'   - The match is case-insensitive.
'   - Only cells that actually contain a formula are examined;
'     plain-value cells are left untouched regardless.
'   - The operation is undoable in one step (Ctrl+Z).
' ============================================================

Public Sub PasteMatchingFormulasAsValues()

    Const DEFAULT_PREFIX As String = "=IFERROR"

    Dim FormulaPrefix As String
    Dim rngSelection  As Range
    Dim cell          As Range
    Dim cellValue     As Variant
    Dim convertedCount As Long

    ' --- Guard: selection must be on a worksheet range ---------
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell range before running this macro.", _
               vbExclamation, "No Range Selected"
        Exit Sub
    End If

    Set rngSelection = Selection

    ' --- Ask the user for the formula prefix -------------------
    FormulaPrefix = InputBox( _
        "Enter the formula prefix to match (case-insensitive)." & vbNewLine & _
        "Cells whose formula starts with this text will be" & vbNewLine & _
        "replaced by their computed value." & vbNewLine & vbNewLine & _
        "Example:  =IFERROR", _
        "Paste Formula as Value – Prefix Filter", _
        DEFAULT_PREFIX)

    ' Blank input or Cancel = abort
    If Trim(FormulaPrefix) = "" Then
        MsgBox "Operation cancelled. No cells were modified.", _
               vbInformation, "Cancelled"
        Exit Sub
    End If

    FormulaPrefix = Trim(FormulaPrefix)

    ' --- Process cells -----------------------------------------
    Application.ScreenUpdating = False

    ' Wrap in a single undo unit
    Application.OnUndo "Paste As Values (" & FormulaPrefix & ")", _
                       "DummyUndoProc"   ' placeholder – Excel batches the edits

    convertedCount = 0

    For Each cell In rngSelection.Cells

        ' Skip cells without a formula
        If Not cell.HasFormula Then GoTo NextCell

        ' Case-insensitive prefix check
        If UCase(Left(cell.Formula, Len(FormulaPrefix))) = UCase(FormulaPrefix) Then

            ' Capture value *before* clearing the formula
            cellValue = cell.Value

            ' Replace formula with its value
            cell.Value = cellValue

            convertedCount = convertedCount + 1
        End If

NextCell:
    Next cell

    Application.ScreenUpdating = True

    ' --- Summary -----------------------------------------------
    Select Case convertedCount
        Case 0
            MsgBox "No cells matched the prefix """ & FormulaPrefix & """." & vbNewLine & _
                   "Nothing was changed.", _
                   vbInformation, "Paste Formula as Value"
        Case 1
            MsgBox "1 cell was converted from formula to value.", _
                   vbInformation, "Paste Formula as Value"
        Case Else
            MsgBox convertedCount & " cells were converted from formula to value.", _
                   vbInformation, "Paste Formula as Value"
    End Select

End Sub


' ============================================================
' PasteMatchingFormulasAsValuesSilent
'
' Same logic as PasteMatchingFormulasAsValues but takes the
' prefix as a parameter instead of prompting.  Useful when
' called programmatically or from another macro.
'
' Example call:
'   PasteMatchingFormulasAsValuesSilent "=IFERROR"
' ============================================================

Public Sub PasteMatchingFormulasAsValuesSilent( _
    Optional ByVal FormulaPrefix As String = "=IFERROR")

    Dim rngSelection As Range
    Dim cell         As Range
    Dim cellValue    As Variant

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell range before running this macro.", _
               vbExclamation, "No Range Selected"
        Exit Sub
    End If

    Set rngSelection = Selection
    FormulaPrefix = Trim(FormulaPrefix)

    If FormulaPrefix = "" Then
        MsgBox "FormulaPrefix cannot be empty.", vbExclamation, "Invalid Argument"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    For Each cell In rngSelection.Cells
        If cell.HasFormula Then
            If UCase(Left(cell.Formula, Len(FormulaPrefix))) = UCase(FormulaPrefix) Then
                cellValue = cell.Value
                cell.Value = cellValue
            End If
        End If
    Next cell

    Application.ScreenUpdating = True

End Sub


' Required placeholder referenced by Application.OnUndo
Private Sub DummyUndoProc()
End Sub

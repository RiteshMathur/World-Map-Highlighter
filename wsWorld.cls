VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsWorld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Dim lookupSheet As String
    Dim rng As Range
    Dim cell As Range
    Dim result As Variant
    
    ' Define the lookup sheet name (modify if needed)
    lookupSheet = wsMapping.Name
    
    ' Set active worksheet
    Set ws = Me
    ws.Unprotect
' Define the range where country selection happens (A3:A40)
    Set rng = ws.Range("A3:A41")
    
    ' Check if the changed cell is within A3:A40
    If Not Intersect(Target, rng) Is Nothing Then
        Application.EnableEvents = False ' Prevent infinite loops
        
        ' Loop through each changed cell
        For Each cell In Target
            If cell.Value = "" Then
                cell.Offset(0, 1).ClearContents ' Clear Column B if deleted
            Else
                cell.Value = Application.WorksheetFunction.Proper(cell.Value)
                result = GetLookupValue(cell.Value, lookupSheet, 2) ' Fetch from function
                cell.Offset(0, 1).Value = result
            End If
        Next cell
        
        Application.EnableEvents = True ' Re-enable events
    End If
    
    
' Define the range where RGB selection happens (C3:C40)
    Set rng = ws.Range("C3:C40")
    
    ' Check if the changed cell is within C3:C40
    If Not Intersect(Target, rng) Is Nothing Then
        Application.EnableEvents = False ' Prevent infinite loops
        
        ' Loop through each changed cell
        For Each cell In Target
            If cell.Value = "" Then
                cell.Offset(0, 1).Interior.ColorIndex = xlNone ' Clear color if empty
            Else
                colorValue = ConvertRGB(cell.Value) ' Convert text to RGB color
                
                If Not IsNull(colorValue) Then
                    cell.Offset(0, 1).Interior.Color = colorValue ' Apply color
                Else
                    cell.Offset(0, 1).Interior.ColorIndex = xlNone ' No color for invalid values
                End If
            End If
        Next cell
        
        Application.EnableEvents = True ' Re-enable events
    End If
    ws.Protect
End Sub


Attribute VB_Name = "Map"

Sub ResetMap()
    Dim iLastRow        As Integer
    Dim rng             As Range
    
    wsWorld.Unprotect
    iLastRow = wsWorld.Cells(wsWorld.Rows.Count, "A").End(xlUp).Row
    Set rng = wsWorld.Range("A3:D" & iLastRow)
    
    If iLastRow > 2 Then
        rng.ClearContents
    End If
    Call ClearMapColor(wsWorld, "World")
    Application.EnableEvents = True
    wsWorld.Protect
End Sub

Function ClearMapColor(ws As Worksheet, strMapName As String)
    Dim objShape        As Object
    Dim objShp          As Object
    
    Set objShape = ws.Shapes(strMapName)
        
    ' Loop through all shapes in the worksheet
    For Each grpShp In ws.Shapes
        ' Check if the shape is a group
        If grpShp.Type = msoGroup Then
            ' Loop through each shape inside the group
            Dim innerShp As Shape
            For Each innerShp In grpShp.GroupItems
                innerShp.Fill.ForeColor.RGB = RGB(165, 165, 165) ' Set color to black
            Next innerShp
        End If
    Next grpShp
End Function

Sub HighlightCountries()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim grpShp As Shape
    Dim innerShp As Shape
    Dim rng As Range
    Dim cell As Range
    Dim countryName As String
    Dim rgbString As String
    Dim rgbParts() As String
    Dim R As Integer, G As Integer, B As Integer
    Dim found As Boolean
    Dim iLastRow As Integer
    
    ' Set the worksheet
    Set ws = ActiveSheet ' Modify if needed
    iLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If iLastRow = 2 Then
        MsgBox "Please update atleast one country name!", vbExclamation, "Map Highlighter"
        Exit Sub
    End If
    
    ' Loop through each selected country in Column A
    For Each cell In ws.Range("B3:B" & iLastRow) ' Adjust range as needed
        countryName = "S_" & Trim(cell.Value) ' Get country name
        
        ' Get the RGB color from Column C
        rgbString = Trim(cell.Offset(0, 1).Value) ' Column C (Offset 2 columns from A)
        
        ' Ensure the RGB string is valid before proceeding
        If Len(rgbString) > 0 And InStr(rgbString, ";") > 0 Then
            ' Split the RGB values
            rgbParts = Split(rgbString, ";")
            If UBound(rgbParts) = 2 Then
                R = Val(Trim(rgbParts(0)))
                G = Val(Trim(rgbParts(1)))
                B = Val(Trim(rgbParts(2)))
            Else
                MsgBox "Invalid RGB format for " & countryName, vbExclamation, "Map Highlighter"
                GoTo NextCountry
            End If
        Else
            MsgBox "Missing or incorrect RGB value for " & countryName, vbExclamation, "Map Highlighter"
            GoTo NextCountry
        End If
        
        found = False
        
        ' Loop through all shapes to find the grouped map
        For Each grpShp In ws.Shapes
            If grpShp.Type = msoGroup Then
                ' Loop through each shape inside the group
                For Each innerShp In grpShp.GroupItems
                    ' Match the country name with the shape name
                    If Trim(innerShp.Name) = countryName Then
                        ' Apply color
                        innerShp.Fill.ForeColor.RGB = RGB(R, G, B)
                        found = True
                    End If
                Next innerShp
            End If
            If found Then Exit For
        Next grpShp
                
NextCountry:
    Next cell

    ' Show message if country shape is not found
    If Not found Then
        MsgBox "Shape not found for country: " & countryName, vbExclamation, "Map Highlighter"
    Else
        MsgBox "Coloring applied successfully! Please verfiy small countries!!!", vbInformation, "Map Highlighter"
    End If

End Sub

Function GetLookupValue(lookupValue As String, lookupSheet As String, lookupCol As Integer) As Variant
    Dim ws As Worksheet
    Dim lookupArr As Variant
    Dim dict As Object
    Dim lastRow As Long, i As Long
    
    ' Set the lookup worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(lookupSheet)
    On Error GoTo 0
    If ws Is Nothing Then
        GetLookupValue = "Sheet Not Found"
        Exit Function
    End If
    
    ' Find last row in lookup sheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        GetLookupValue = "No Data"
        Exit Function
    End If
    
    ' Load lookup data into an array
    lookupArr = ws.Range("B2:C" & lastRow).Value
    
    ' Store lookup data in a dictionary
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(lookupArr, 1)
        dict(lookupArr(i, 1)) = lookupArr(i, lookupCol)
    Next i
    
    ' Return the matched value or "Not Found"
    If dict.exists(lookupValue) Then
        GetLookupValue = dict(lookupValue)
    Else
        GetLookupValue = "Not Found"
    End If
End Function

Function ConvertRGB(rgbText As String) As Variant
    Dim rgbParts As Variant
    Dim R As Integer, G As Integer, B As Integer
    
    ' Split RGB format (e.g., "255; 0; 0" ? [255, 0, 0])
    rgbParts = Split(Trim(rgbText), ";")
    
    ' Check if valid RGB format
    If UBound(rgbParts) = 2 Then
        On Error Resume Next
        R = CInt(Trim(rgbParts(0)))
        G = CInt(Trim(rgbParts(1)))
        B = CInt(Trim(rgbParts(2)))
        On Error GoTo 0
        
        ' Validate RGB values (0-255)
        If R >= 0 And R <= 255 And G >= 0 And G <= 255 And B >= 0 And B <= 255 Then
            ConvertRGB = RGB(R, G, B)
            Exit Function
        End If
    End If
    
    ' Return Null for invalid RGB formats
    ConvertRGB = Null
End Function


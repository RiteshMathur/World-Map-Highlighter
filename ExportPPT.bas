Attribute VB_Name = "ExportPPT"
Option Explicit

Function ExportMapToPowerPoint(sheetName As String, Optional mapName As String = "") As Boolean
    Dim pptApp As Object ' PowerPoint Application
    Dim pptPres As Object ' PowerPoint Presentation
    Dim pptSlide As Object ' PowerPoint Slide
    Dim ws As Worksheet
    Dim mapShape As Shape
    Dim slideWidth As Single, slideHeight As Single
    Dim shapeFound As Boolean

    ' Define slide dimensions for 4:3 aspect ratio
    slideWidth = 1024 ' Adjust if needed
    slideHeight = 768 ' Adjust if needed

    ' Error Handling
    On Error GoTo ErrorHandler

    ' Set worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' Find the map shape
    shapeFound = False
    For Each mapShape In ws.Shapes
        If mapName = "" Or mapShape.Name = mapName Then ' Match name if provided
            shapeFound = True
            mapShape.Copy
            Exit For
        End If
    Next mapShape

    ' If no shape found, exit with error
    If Not shapeFound Then
        MsgBox "Map not found! Ensure the shape exists in '" & sheetName & "'.", vbExclamation, "Error"
        ExportMapToPowerPoint = False
        Exit Function
    End If

    ' Create a new PowerPoint application
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True ' Make PowerPoint visible

    ' Create a new presentation
    Set pptPres = pptApp.Presentations.Add

    ' Set slide size to 4:3
    pptPres.PageSetup.slideWidth = slideWidth
    pptPres.PageSetup.slideHeight = slideHeight

    ' Add a new slide
    Set pptSlide = pptPres.Slides.Add(1, 12) ' ppLayoutText = 1 (Title Slide)

    ' Paste the copied map onto the slide
    pptSlide.Shapes.PasteSpecial DataType:=2 ' ppPasteEnhancedMetafile

    ' Resize and center the pasted shape
    With pptSlide.Shapes(pptSlide.Shapes.Count)
        .LockAspectRatio = msoTrue
        .Left = (slideWidth - .Width) / 2
        .Top = (slideHeight - .Height) / 2
    End With

    ' Success message
    ExportMapToPowerPoint = True
    Exit Function

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    ExportMapToPowerPoint = False
End Function

Sub ExportSpecificMap()
    If ExportMapToPowerPoint(wsWorld.Name, "World") Then
        MsgBox "World map exported successfully!", vbInformation, "Success"
    End If
End Sub



Sub SplitTextIntoObjects()
    Dim shp As Visio.Shape
    Dim text As String
    Dim textArray() As String
    Dim i As Integer
    Dim newShape As Visio.Shape
    Dim xPos As Double
    Dim yPos As Double
    Dim lineHeight As Double
    Dim textWidth As Double

    ' Check if there is a selected shape
    If Visio.ActiveWindow.Selection.Count = 0 Then
        MsgBox "Please select a shape with multi-line text."
        Exit Sub
    End If
    
    Set shp = Visio.ActiveWindow.Selection.PrimaryItem
    
    ' Get the shape's text
    text = shp.text
    
    ' Split the shape's text into lines using both CR and LF
    textArray = Split(Replace(text, vbCrLf, vbLf), vbLf)
    
    ' Debug: Show the number of lines detected
    MsgBox "Number of lines detected: " & UBound(textArray) - LBound(textArray) + 1
    
    ' Get the position and width of the original shape
    xPos = shp.CellsU("PinX").ResultIU
    yPos = shp.CellsU("PinY").ResultIU
    textWidth = shp.CellsU("Width").ResultIU
    
    ' Debug: Show the original shape position and width
    MsgBox "Original shape position (X, Y): (" & xPos & ", " & yPos & ")" & vbCrLf & "Width: " & textWidth
    
    ' Set the height for each line of text
    lineHeight = shp.CellsU("Height").ResultIU / (UBound(textArray) - LBound(textArray) + 1)
    
    ' Loop through each line of text and create a new text shape for each line
    For i = LBound(textArray) To UBound(textArray)
        Set newShape = Visio.ActivePage.DrawRectangle(xPos - textWidth / 2, yPos - (i * lineHeight), xPos + textWidth / 2, yPos - ((i + 1) * lineHeight) + lineHeight)
        newShape.text = textArray(i)
        newShape.CellsU("Height").ResultIU = lineHeight
        newShape.CellsU("Width").ResultIU = textWidth
        
        ' Debug: Show the text of the new shape
        MsgBox "Created shape with text: " & textArray(i)
    Next i
    
    ' Delete the original shape
    shp.Delete
    
    ' Debug: Confirm deletion
    MsgBox "Original shape deleted."
End Sub

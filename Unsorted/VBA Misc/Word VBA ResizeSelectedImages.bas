Sub ResizeSelectedImages()
    Dim userInput As String
    Dim targetWidthCm As Double
    Dim shp As Shape
    Dim ils As InlineShape
    Dim aspectRatio As Double

    ' Prompt user for width with default value
    userInput = InputBox("Enter desired image width in cm:", "Resize Selected Images", "7.93")
    
    If userInput = "" Then Exit Sub ' User cancelled
    If Not IsNumeric(userInput) Then
        MsgBox "Please enter a valid number.", vbExclamation
        Exit Sub
    End If
    
    targetWidthCm = Val(userInput)
    
    ' Convert cm to points (1 cm = 28.35 points)
    Dim targetWidthPts As Double
    targetWidthPts = targetWidthCm * 28.35

    ' Resize InlineShapes
    For Each ils In Selection.InlineShapes
        If ils.Type = wdInlineShapePicture Or ils.Type = wdInlineShapeLinkedPicture Then
            aspectRatio = ils.Height / ils.Width
            ils.LockAspectRatio = msoFalse
            ils.Width = targetWidthPts
            ils.Height = targetWidthPts * aspectRatio
        End If
    Next ils

    ' Resize Shapes
    For Each shp In Selection.ShapeRange
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            aspectRatio = shp.Height / shp.Width
            shp.LockAspectRatio = msoFalse
            shp.Width = targetWidthPts
            shp.Height = targetWidthPts * aspectRatio
        End If
    Next shp

    MsgBox "Selected images resized to " & targetWidthCm & " cm wide.", vbInformation
End Sub
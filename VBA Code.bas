Attribute VB_Name = "Module1"
Sub AddImage()
    
    Dim product As String
    Dim rngURL As Range
    Dim strURL As String
    Dim thePix As Shape
    Dim iCC As Long
    Dim style As String
    Dim blnFail As Boolean
    Dim wks As Worksheet
    
    Set rngURL = ActiveCell
    product = Left(ActiveCell, 8)
    iCC = 0

    If IsNumeric(product) And Len(product) = 8 Then
        strURL = "http://oldnavy.gap.com/resources/productImage/v1/" & product & "2/P01"
        Else: strURL = "http://oldnavy.gap.com/resources/productImage/v1/" & Left(product, 6) & "0" & iCC & "2/P01"
        style = "Yes"
    End If
    rngURL.Select
    Set wks = ActiveSheet
' Import image; first insert rectangle shape, then fill it with the image
    Set thePix = wks.Shapes.AddShape(msoShapeRectangle, rngURL.Left, _
        rngURL.Top, 120, 140)
    Application.DisplayAlerts = False
    On Error GoTo NO_IMAGE
    thePix.Fill.UserPicture strURL
    With thePix
        .Top = rngURL.Top
        .Left = rngURL.Left
    End With
    Application.DisplayAlerts = True
' Size image
    imgTargetHeight = 140
    imgTargetWidth = 120
    imgWidth = thePix.Width
    imgHeight = thePix.Height
    If imgHeight >= imgWidth Then
        thePix.LockAspectRatio = msoTrue
        thePix.Height = imgTargetHeight
    Else
        thePix.LockAspectRatio = msoTrue
        thePix.Width = imgTargetWidth
    End If
' Position image
    imgContainerWidth = 180
    imgWidth = thePix.Width
    imgMove = (imgContainerWidth - imgWidth) / 2
    thePix.IncrementLeft imgMove
    imgContainerHeight = 140
    imgHeight = thePix.Height
    imgMoveDown = (imgContainerHeight - imgHeight) / 2
    thePix.IncrementTop imgMoveDown
    Set thePix = Nothing
' Return nothing
    Exit Sub
    
NO_IMAGE:
'attempted to download image for specific cc. If that fails, tag foundTab
'as false. This tells parent procedure to continue trying to find the image
'of any cc in the style (-00 through -09) and use it rather than nothing.
    
    thePix.Delete
    Set thePix = Nothing
    If style = "Yes" And iCC < 7 Then
        iCC = iCC + 1
            If IsNumeric(product) And Len(product) = 8 Then
        strURL = "http://oldnavy.gap.com/resources/productImage/v1/" & product & "2/P01"
        Else: strURL = "http://oldnavy.gap.com/resources/productImage/v1/" & Left(product, 6) & "0" & iCC & "2/P01"
        style = "Yes"
    End If
    rngURL.Select
    Set wks = ActiveSheet
' Import image; first insert rectangle shape, then fill it with the image
    Set thePix = wks.Shapes.AddShape(msoShapeRectangle, rngURL.Left, _
        rngURL.Top, 120, 140)
    Application.DisplayAlerts = False
    Resume
    End If
    MsgBox ("Image not found")
    Application.DisplayAlerts = True
End Sub



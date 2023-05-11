Sub Resize_SitePhotos()
'
' Resize Macro
'
'
    Dim pic As InlineShape
   
    For Each pic In ActiveDocument.Section(4).Range.InlineShapes
       
        With pic
            .LockAspectRatio = msoTrue
            Xw = .Width
            Xh = .Height
            Y = 17.5
            ' If Xw > Xh Then ' horizontal
                .Height = CentimetersToPoints(6)
           
                ' .Height = Y * Xh / Xw
            ' Else  ' vertical
            '    .Height = CentimetersToPoints(Y)
                ' .Width = CentimetersToPoints(Y * Xw / Xh)
                ' .Width = Xh * Y / Xw
            ' End If
        End With
    Next

    Dim pShape As Word.InlineShape

    For Each pShape In ActiveDocument.Section(4).Range.InlineShapes
        With pShape.Range
            .InsertAfter vbTab
        End With

    Next
End Sub00

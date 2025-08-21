Attribute VB_Name = "Module1"
Sub StripAllDesignCustomisationsWithMinFontSizeAndRemoveEmptyTextboxes()
    Dim sld As Slide
    Dim shp As Shape
    Dim txtRng As TextRange
    Dim i As Long

    For Each sld In ActivePresentation.Slides
        sld.FollowMasterBackground = msoFalse

        sld.CustomLayout = ActivePresentation.Designs(1).SlideMaster.CustomLayouts(1)

        With sld.Background.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(200, 200, 200)
            .BackColor.RGB = RGB(200, 200, 200)
            .Solid
        End With

        For i = sld.Shapes.Count To 1 Step -1
            With sld.Shapes(i)
                If .Type = msoPicture Then
                    .Delete
                ElseIf .Fill.Type = msoFillPicture Or .Fill.Type = msoFillGradient Then
                    .Delete
                End If
            End With
        Next i

        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    Set txtRng = shp.TextFrame.TextRange
                    With txtRng.Font
                        .Color.RGB = RGB(0, 0, 0)
                        .Bold = msoFalse
                        .Italic = msoFalse
                        .Shadow = msoFalse
                        .Underline = msoFalse
                        .Name = "Arial"
                        
                        If txtRng.Font.Size < 18 Then
                            .Size = 18
                        End If
                    End With
                End If
            End If
        Next shp

        For i = sld.Shapes.Count To 1 Step -1
            Set shp = sld.Shapes(i)
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                Else
                    shp.Delete
                End If
            End If
        Next i
    Next sld

    MsgBox "Powerpoint greyed out"
End Sub




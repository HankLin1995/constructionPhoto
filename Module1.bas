Attribute VB_Name = "Module1"
Sub ⅷ떠1()
Attribute ⅷ떠1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ⅷ떠1 ⅷ떠
'

'
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "123"
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 3).ParagraphFormat. _
        FirstLineIndent = 0
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 3).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 11
        .Name = "+mn-lt"
    End With
    Range("E6").Select
    ActiveSheet.Shapes.Range(Array("TextBox 260")).Select
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(146, 208, 80)
        .Transparency = 0
        .Solid
    End With
End Sub

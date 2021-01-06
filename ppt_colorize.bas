Colors = Array(RGB(255, 0, 0), RGB(0, 255, 0), RGB(0, 0, 255), RGB(255, 255, 0), RGB(255, 0, 255), RGB(0, 255, 255))
Set tr = ActivePresentation.Slides(1).Shapes(1).TextFrame.TextRange
For i = 1 To tr.Words.Length
  tr.Words(i, Length:=1).Font.color.RGB = Colors((i - 1) Mod 6)
Next

Sub GetMyNodePosition()
    myFile = "C:\data\MP3.csv"
    Open myFile For Output As #1
    Dim s As Shape, n As Node
    Dim sh As Shape
    Dim x As Double, y As Double
    Dim x1 As Double, y1 As Double
    Dim x2 As Double, y2 As Double
    Dim x3 As Double, y3 As Double
    Dim x4 As Double, y4 As Double
    Dim colr As Double
    
    Set s = ActiveShape

    If s.Type = cdrCurveShape Then
        For Each n In s.Curve.Nodes
            n.GetPosition x, y
            Write #1, x,
            Write #1, y
        Next n
    End If
    Close #1
End Sub

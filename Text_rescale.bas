Attribute VB_Name = "Module1"
Sub rescale_and_save_fonts()
'macros can return scale of your text in a page to an original proportions.
Dim CentX_original As Double
Dim CentY_original As Double

Dim sr As ShapeRange
Dim s As Shape

'define the scaling percentage
    vert_scale_pctg = InputBox("Vertical rescale in percentages you applied to your picture")
    hor_scale_pctg = InputBox("Horizontal rescale in percentages you applied to your picture")
'set shape range to the active selection
Set sr = ActiveSelectionRange

'if no shapes are selected, then notify user and exit the sub
If sr.Count = 0 Then
    MsgBox "Nothing is selected!"
    Exit Sub
End If
   
'work through all of the shapes in the shape range
For Each s In sr
If s.Type = cdrTextShape Then
'remember the Center X and Center Y positions for this shape
CentX_original = s.CenterX
CentY_original = s.CenterY


'apply the desired scaling to this shape, based on *original* size
s.SizeHeight = s.SizeHeight * 100 / vert_scale_pctg
s.SizeWidth = s.SizeWidth * 100 / hor_scale_pctg
'reposition this shape to its original Center X and Center Y positions
s.CenterX = CentX_original
s.CenterY = CentY_original

End If
Next s

End Sub

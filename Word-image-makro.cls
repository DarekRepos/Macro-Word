VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Makro pozwalla zrobiæ Obrazki na wys. 9.cm, zwiêkszyæ/zmniekszyc wysokoœæ i szerokosc obrazka o 1 cm


Option Explicit

Sub SetMultipleImagesSize()

Dim i As Long

With ActiveDocument
    For i = 1 To .InlineShapes.Count
        With .InlineShapes(i)
             .ScaleHeight = 23
             .ScaleWidth = 23
        End With
    Next i
End With

End Sub


Sub IncreaseImagesSize()

Dim i As Long
Dim mysize As Long
Dim CurrentSize As Long

mysize = 3

If Selection.Type <> wdSelectionIP Then
    For i = 1 To Selection.InlineShapes.Count
        With Selection.InlineShapes(i)
            CurrentSize = .ScaleHeight
            .ScaleHeight = CurrentSize + mysize
            .ScaleWidth = CurrentSize + mysize
        End With
    Next i
Else
 MsgBox Prompt:="Not a valid selection! Exiting procedure..."
 Exit Sub
End If

End Sub


Sub DecreaseImagesSize()

Dim i As Long
Dim mysize As Long
Dim CurrentSize As Long

mysize = 3

If Selection.Type <> wdSelectionIP Then
    For i = 1 To Selection.InlineShapes.Count
        With Selection.InlineShapes(i)
            CurrentSize = .ScaleHeight
            .ScaleHeight = CurrentSize - mysize
            .ScaleWidth = CurrentSize - mysize
        End With
    Next i
Else
 MsgBox Prompt:="Not a valid selection! Exiting procedure..."
 Exit Sub
End If

End Sub



Dim swApp As SldWorks.SldWorks
Dim swDraw As SldWorks.DrawingDoc
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

Sub main()

Set swApp = Application.SldWorks
Set swDraw = swApp.ActiveDoc

DimAdjust

End Sub

Sub DimAdjust()

Set swApp = Application.SldWorks
Set swModel = swApp.ActiveDoc
Dim swDraw As SldWorks.DrawingDoc
Set swDraw = swModel
Dim v As View
Set v = swDraw.GetFirstView 'sheet
Set v = v.GetNextView 'first drawing view

Do While Not v Is Nothing

Dim c As Integer

c = v.GetAnnotationCount

Dim allAnnotations As Variant
allAnnotations = v.GetAnnotations

Dim i As Integer

For i = 0 To c - 1

Dim ann As Annotation
Set ann = allAnnotations(i)

If ann.GetType = swDisplayDimension Then 'found a dimension

    Dim swDim As DisplayDimension
    Set swDim = ann.GetSpecificAnnotation
    Dim swDimType As Long
    swDimType = swDim.Type2 'get the dimension type
    Dim swDimName As String
    swDimName = swDim.GetNameForSelection 'get the dimension name
    Dim pos() As Double
    pos = ann.GetPosition 'get it's position on the sheet
    Dim posX As Double
    Dim posY As Double
    posX = pos(0)
    posY = pos(1)
    Dim res As Boolean
    res = ann.Select(True)
    
End If
        
boolstatus = swDraw.Extension.AlignDimensions(0, 0.175)
    
Next i
Set v = v.GetNextView
Loop

Stop

End Sub


Option Explicit

Sub main()

Dim swApp                   As SldWorks.SldWorks
Dim swModel                 As SldWorks.ModelDoc2
Dim vConfNameArr            As Variant
Dim sConfigName             As String
Dim nStart                  As Single
Dim i                       As Long
Dim j                       As Integer
Dim k                       As Integer
Dim arr()                   As String
Dim bShowConfig             As Boolean
Dim bRebuild                As Boolean
Dim bRet                    As Boolean
Dim CurFeature              As SldWorks.Feature

Set swApp = CreateObject("SldWorks.Application")
Set swModel = swApp.ActiveDoc

' Is a document active
If swModel Is Nothing Then
    swApp.SendMsgToUser2 "A sheet metal part must be open.", swMbWarning, swMbOk
    Exit Sub
End If

' Is it a part model
Dim modelType As Long
modelType = swModel.GetType
If modelType <> SwConst.swDocPART Then
    swApp.SendMsgToUser2 "A sheet metal part must be open.", swMbWarning, swMbOk
    Exit Sub
End If

'   List containing all configuration names
vConfNameArr = swModel.GetConfigurationNames

For i = 0 To UBound(vConfNameArr)
'   Skips anything from happening to flat patterns
    sConfigName = vConfNameArr(i)
    If InStr(sConfigName, "FLAT-PATTERN") Then
        GoTo Skip:
    End If
    bShowConfig = swModel.ShowConfiguration2(sConfigName)
    bRebuild = swModel.ForceRebuild3(False)
    
    Dim FilePath As String
    FilePath = swModel.GetPathName
    arr = Split(FilePath, "\")
    j = Len(arr(7)) - 7
    k = Len(arr(6)) + 1
    
    FilePath = VBA.Left(FilePath, Len(FilePath) - 7) 'Take off solidworks file extension
    FilePath = VBA.Left(FilePath, Len(FilePath) - j) 'Take off current configuration ID
    FilePath = VBA.Left(FilePath, Len(FilePath) - k) 'Take off file extension and replace with desired
    
    FilePath = FilePath & "Library Generation Related\Macro-Generated Drawings\" & sConfigName

    bRet = swModel.ExportFlatPatternView(FilePath & ".DXF", 1)
Skip:
Next i

MsgBox "Completed"

End Sub


Dim swApp As SldWorks.SldWorks
Dim swDoc As SldWorks.ModelDoc2
Dim swFeatMgr As SldWorks.FeatureManager
Sub main()

Set swApp = Application.SldWorks
Set swDoc = swApp.GetFirstDocument
While Not swDoc Is Nothing
Set swFeatMgr = swDoc.FeatureManager
swFeatMgr.EditRollback swMoveRollbackBarToEnd, ""
Set swDoc = swDoc.GetNext
Wend
End Sub

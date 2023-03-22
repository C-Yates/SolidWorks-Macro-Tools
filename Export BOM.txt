' This macro creates an xls file from either a SolidWorks type BOM which is hard coded
' or
' by the user selecting a SolidWorks type BOM in the feature manager design tree

' Preconditions that a drawing is open and contains a SolidWorks type BOM

' You need to add references for:
' SolidWorks 20xx Type Library
' SolidWorks 20xx Constant Type Library
' Microsoft Excel XX.0 Object Libary

' Change the xx part above to the year you have on your system

    Option Explicit


Sub main()

    On Error GoTo ErrH:
    
    Dim swApp         As SldWorks.SldWorks
    
    Dim swModelDoc    As SldWorks.ModelDoc2
    
    Dim swSelMgr      As SldWorks.SelectionMgr
    
    Dim swTableAnn    As SldWorks.TableAnnotation

    Dim swBomFeature  As SldWorks.BomFeature
    
    Dim swAnn         As SldWorks.Annotation
    
    Dim vTableArr     As Variant
    
    Dim vTable        As Variant
    
    Dim retval        As Boolean
    
    Dim CSVFile       As String
    
    
    
    Set swApp = Application.SldWorks

    Set swModelDoc = swApp.ActiveDoc
    
    Set swSelMgr = swModelDoc.SelectionManager
        
    ' You can either run through the feature manager design tree and hard code in the name of a BOM
    ' So the user does not have to select a BOM evrytime
    ' Go to this function (TraverseFeatureTree) and change the name of the BOM
    ' This function will run through the feature tree and find a BOM
    ' Comment out the line below if you want to manually select a BOM in the feature tree
    TraverseFeatureTree
    
    
    ' Make sure a BOM is selected in the feature manager design tree
    Set swBomFeature = swSelMgr.GetSelectedObject5(1)
    
    
    ' Make sure a BOM is selected in the feature manager design tree
    If swBomFeature Is Nothing Then
    
        MsgBox "Please select a BOM to export"
        Exit Sub
    
    End If
    
    
    vTableArr = swBomFeature.GetTableAnnotations
        
    For Each vTable In vTableArr
        
        ' Got BOM as table annotation
        Set swTableAnn = vTable

    Next vTable
        
    ' Rename BOM with .csv file extension
    CSVFile = RenameBomToCSV
    
    
    ' Save csv file. If you save it as an xls file and try to open it in Excel and
    ' it will tell you that it is an text file.
    ' This way it actually saves as a csv file and no message box will pop up
    retval = swTableAnn.SaveAsText(CSVFile, ",")
    
    ' Now change file extension to .xls and save
    SaveCSVAsXLS CSVFile
    
    ' Get rid of .csv file
    DeleteFile (CSVFile)
    
    ' Complete process
    MsgBox "BOM processed"
    
    
    ' Clean up
    Set swBomFeature = Nothing
        
    Set swModelDoc = Nothing
        
    Set swApp = Nothing
    
    
ErrH:
    
    If Err.Number = 0 Or Err.Number = 20 Then
    
        Resume Next
        
    Else
    
        ' Type mismatch
        If swBomFeature Is Nothing Then
    
            MsgBox "Please select a BOM from the Feature Manager Tree"
            Exit Sub
    
        Else
    
            MsgBox Err.Number & " " & Err.Description
        
        End If
    
    End If
    
End Sub

Sub TraverseFeatureTree() ' You could even add arguments

    
    Dim swApp As SldWorks.SldWorks
    
    Dim swModelDoc As SldWorks.ModelDoc2
    
    Dim swFeature As SldWorks.Feature
    
    Dim ModelDocType As Long
    
    Dim FeatureName As String
    
    
    ' Connect to SW
    Set swApp = Application.SldWorks
    
    ' Get active document
    Set swModelDoc = swApp.ActiveDoc
    
    ' Clear any selection
    swModelDoc.ClearSelection
    
    ' Get document type
    ModelDocType = swModelDoc.GetType
    
    ' Get first feature in feature tree
    Set swFeature = swModelDoc.FirstFeature
    
    
        ' Start traversal
        While Not swFeature Is Nothing

            FeatureName = swFeature.Name
            
            Debug.Print FeatureName
                
                ' Do what you want here. I just searched the feature tree for a BOM called Bill of Materials2
                ' Change "Bill of Materials2" to the BOM of your choice
                If Left(FeatureName, 17) = "Bill of Materials" Then
            
                    ' Select the BOM
                    swFeature.Select True
                    
                    ' Exit early
                    Exit Sub
                
                End If
        
            ' Get next feature
            Set swFeature = swFeature.GetNextFeature
        
        Wend
    
End Sub

Function RenameBomToCSV() As String
      
    Dim swApp         As SldWorks.SldWorks
    
    Dim swModelDoc    As SldWorks.ModelDoc2
    
    Dim GetPath       As String
    
    'clear string
    RenameBomToCSV = ""
    
    Set swApp = Application.SldWorks

    Set swModelDoc = swApp.ActiveDoc
    
    'Get full path of active document
    GetPath = swModelDoc.GetPathName
    
    'take off solidworks file extension
    GetPath = VBA.Left(GetPath, Len(GetPath) - 6)
    
    'now add csv file extension
    GetPath = GetPath & "csv"
    
    RenameBomToCSV = GetPath
    
    'clean up
    Set swModelDoc = Nothing
        
    Set swApp = Nothing

End Function

' Pass in the CSV file
Sub SaveCSVAsXLS(WhichDoc As String)
    
    Dim xlApp As Excel.Application

    Dim xlWB  As Excel.Workbook

    Dim FileToKill As String
    
    ' If there is an existing file it will get deleted
    FileToKill = VBA.Left(WhichDoc, Len(WhichDoc) - 3) & "xls"
    
    Debug.Print FileToKill
    
    
    If Dir(FileToKill) <> "" Then
    
        ' Kill the existing file to stop a message popping up
        ' File already exists do you want to replace it
        ' This just make it a bit slicker
        Kill FileToKill
        
        Set xlApp = CreateObject("Excel.Application")
    
        xlApp.Visible = False
        
        ' Open the CSV file
        Set xlWB = xlApp.Workbooks.Open(WhichDoc)
        
        ' and save as xls
        xlWB.SaveAs VBA.Left(WhichDoc, Len(WhichDoc) - 3) & "xls", 56
        
        ' Show the xls file
        xlApp.Visible = True
        
    Else
        
        Set xlApp = CreateObject("Excel.Application")
    
        xlApp.Visible = False

        Set xlWB = xlApp.Workbooks.Open(WhichDoc)
    
        xlWB.SaveAs VBA.Left(WhichDoc, Len(WhichDoc) - 3) & "xls", 56
        
        xlApp.Visible = True
    
    End If
    
End Sub

Sub DeleteFile(DeleteWhichFile As String)

    Kill DeleteWhichFile

End Sub

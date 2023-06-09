' Credit: Luke Malpass
' Website: http://www.angelsix.com
'
' Prerequisits: Have a drawing open, and for it to export anything, have at least one table created (such as a BOM)
'
'

Option Explicit

Dim swApp                   As SldWorks.SldWorks
Dim swModel                 As SldWorks.ModelDoc2
Dim swDraw                  As SldWorks.DrawingDoc
Dim swView                  As SldWorks.View

Sub main()

    Dim swTable                 As SldWorks.TableAnnotation
    Dim bRet                    As Boolean

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swDraw = swModel

    ' Get the first view
    Set swView = swDraw.GetFirstView

    Do While Not swView Is Nothing
        ' Show the name of the view
        Debug.Print "  " & swView.Name
        
        ' Get the first table annotation for this view
        Set swTable = swView.GetFirstTableAnnotation

        Do While Not swTable Is Nothing
            ' Use built in function to manually export BOM to text file, but will open fine as XLS with warning
            ' that is is a text file. This function has a few drawbacks...
            swTable.SaveAsText "c:\Standard BOM.xls", vbTab
            ' So create our own manual method instead
            ExportTable swTable
            
            ' Get next table annotation for this view
            Set swTable = swTable.GetNext
        Loop
        
        Set swView = swView.GetNextView
    Loop

End Sub

Sub ExportTable(swTable As SldWorks.TableAnnotation)

    Dim swAnn                   As SldWorks.Annotation
    Dim nNumCol                 As Long
    Dim nNumRow                 As Long
    Dim sRowStr                 As String
    Dim i                       As Long
    Dim j                       As Long
    
    Dim objExcelObject          As Excel.Application
    Dim objBook1                As Excel.Workbook
    Dim objSheet1               As Excel.Worksheet
    
    Set objExcelObject = CreateObject("Excel.Application")
    objExcelObject.Visible = False
    objExcelObject.ScreenUpdating = False
    objExcelObject.Interactive = False
    
    Set objBook1 = objExcelObject.Workbooks.Add()
    Set objSheet1 = objBook1.ActiveSheet
        
    Set swAnn = swTable.GetAnnotation
    
    objBook1.SaveAs "c:\" & swAnn.GetName

    objSheet1.Activate

    'Set worksheet name
    objSheet1.Name = "BOM"


    nNumCol = swTable.ColumnCount
    nNumRow = swTable.RowCount

    ' Only export Bills of Materials (EDIT, OR REMOVE TO PROCESS ALL TABLES)
    If swTable.Type <> SwConst.swTableAnnotationType_e.swTableAnnotation_BillOfMaterials Then Exit Sub
    
    ' Get the table contents
    For i = 0 To nNumRow - 1
        For j = 0 To nNumCol - 1
            objSheet1.Range("A1").Offset(i, j).Value = swTable.DisplayedText(i, j)
            ' Bold top column
            If i = 0 Then objSheet1.Range("A1").Offset(i, j).Font.Bold = True
        Next j
    Next i
        
    objSheet1.Range("A1").Activate
    objSheet1.Range("A1", objSheet1.Range("A1").SpecialCells(xlCellTypeLastCell)).Columns.AutoFit
        
    ' Clean-up
    objExcelObject.ScreenUpdating = True
    objExcelObject.Interactive = True
    
    objExcelObject.ActiveWorkbook.Save
    objExcelObject.Workbooks.Close
    objExcelObject.Quit
    
    Set objSheet1 = Nothing
    Set objBook1 = Nothing
    Set objExcelObject = Nothing
    
    
End Sub


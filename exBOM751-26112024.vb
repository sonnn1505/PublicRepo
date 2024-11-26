'class1

'Public PartNum     As String
'Public Qty         As Long
'Public swModel     As SldWorks.ModelDoc2
'Public PathFile    As String
'Public Description As String

'Public category    As String
'Public chunggLoai  As String
'Public DrawingNo   As String
'Public Material    As String
'Public Weight      As String
'Public Length      As String
'Public HeatTreatment As String
'Public SurfaceProtection As String
'Public SurfaceFinish As String
'Public Comment     As String
'Public Lable       As String
'public isWelment   As Boolean
'public myParent    As String

Type BomPosition
    indent          As String
    ModelPath       As String
    Configuration   As String
    Quantity        As Double
    Price           As Double
    category        As String
    chunggLoai      As String
    drawingNo       As String
    Description     As String
    Material        As String
    Weight          As String
    partNumber      As String
    Length          As String
    HeatTreatment   As String
    SurfaceProtection As String
    SurfaceFinish   As String
    Comment         As String
    isWelment       As Boolean
    PartNumParentWelment As String
    isParentOfWelment As Boolean
    isSLDASM        As Boolean
    isDirect        As Boolean
    myParent        As String
    
End Type
Dim fso             As Object

Dim MyItems         As Collection
Dim item            As Class1

Const FILE_PATH     As String = "C:\Workspaces\Template\Macro Soluca\751_BOM_SX.xlsx"
Const THUMNAIL_PATH As String = "C:\Workspaces\"
Const THUMNAIL_PATH_WELMENT As String = "C:\Workspaces\Weldment Profile\"
Dim exApp           As Object
Dim exWorkbook      As Object
Dim exWorkSheet     As Object
Dim swSelectionManager As Object
'define the width and height of the thumbnail
Dim Width           As Long        'in pixels
Dim Height          As Long        'in pixels
Dim RowStart        As Long
Dim Header          As String
Dim thumbnailPath   As String

Dim swApp           As SldWorks.SldWorks

Sub Main()
    
    Width = 21
    Height = 60
    
    RowStart = 10
    Set swApp = Application.SldWorks
    
    Dim swAssy      As SldWorks.AssemblyDoc
    
    Set swAssy = swApp.ActiveDoc
    
    If Not swAssy Is Nothing Then
        
        swAssy.ResolveAllLightWeightComponents True
        
        'Get Header file cum tong
        Dim swRootComp As SldWorks.Component2
        Set swRootComp = swAssy.ConfigurationManager.ActiveConfiguration.GetRootComponent
        Dim swCompModel As SldWorks.ModelDoc2
        Set swCompModel = swRootComp.GetModelDoc2()
        Dim Desc    As String
        
        'root properties
        Dim activeConfigNames  As String
        'Get and print names of all configurations
        activeConfigNames = swApp.GetActiveConfigurationName(swRootComp.GetPathName())
        
		' remove this condition 26/11/2024
		'If activeConfigNames <> "Default" Then
        '    swApp.SendMsgToUser "Please active your config TAB <Default> to Export BOM."
        '    Exit Sub
        'End If
        
        Dim filePath    As String
        filePath = swRootComp.GetPathName()
        Dim PartNameFromPath As String
        PartNameFromPath = Mid(filePath, InStrRev(filePath, "\") + 1, InStrRev(filePath, ".") - InStrRev(filePath, "\") - 1)
    
        
        Desc = GetPropertyValue(swCompModel, swRootComp.ReferencedConfiguration, "Description", PartNameFromPath)
        Dim Number  As String
        Number = GetPropertyValue(swCompModel, swRootComp.ReferencedConfiguration, "Number", PartNameFromPath)
        Header = Desc & "_" & Number
        
        'Get foder of file cum
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim Folder  As String
        Folder = Dir(fso.GetParentFolderName(swRootComp.GetPathName()), vbDirectory)
        thumbnailPath = THUMNAIL_PATH & Folder & "\"
        'clean up
        Set fso = Nothing
        'tree bom
        Dim swConfig  As SldWorks.Configuration
        Dim swConfigMgr As SldWorks.ConfigurationManager
        Dim swRootComp3 As SldWorks.Component2
        Set swConfigMgr = swAssy.ConfigurationManager
        Set swConfig = swConfigMgr.ActiveConfiguration
        Set swRootComp3 = swConfig.GetRootComponent3(True)
        
        'Get Tree BOM ITEM by collection
        Set MyItems = New Collection
        SetCompVisib swRootComp, 1, "1", ""
        
        'Get Flat BOM
        Dim bom()   As BomPosition
        bom = GetFlatBom(swAssy)

        'map welment to tree BomPosition
        'map welment to tree BomPosition
        Dim temp    As Collection
        Set temp = New Collection
        For Each item In MyItems
            temp.Add item
            
            If item.PathFile Like "*.SLDPRT" Then
                Dim ParTR As String
                ParTR = item.PartNum
                
                'find in Flat Bom
                Dim prefixWel As Integer
                prefixWel = 1
                Dim fl As Integer
                For fl = 0 To UBound(bom)
                    If bom(fl).isWelment Then
                        Dim parFL As String
                        parFL = bom(fl).PartNumParentWelment
                        If ParTR = parFL Then        'And bom(fl).partNumber <> ""
                        
                        Dim itemWEL As Class1
                        Set itemWEL = New Class1
                        itemWEL.Qty = bom(fl).Quantity
                        itemWEL.Description = bom(fl).Description
                        itemWEL.PartNum = bom(fl).partNumber
                        itemWEL.Length = bom(fl).Length
                        itemWEL.Material = bom(fl).Material
                        itemWEL.drawingNo = bom(fl).drawingNo
                        itemWEL.Lable = item.Lable & "." & prefixWel
                        If itemWEL.PartNum Like "00*" Then
                            itemWEL.category = "SPTC (San Pham Tieu Chuan)"
                        Else
                            itemWEL.category = "VTGC (Vat Tu Gia Cong)"
                        End If
                        itemWEL.isWelment = True
                        
                        temp.Add itemWEL
                        
                        prefixWel = prefixWel + 1
                    End If
                End If
                
            Next fl
        End If
    Next
    
    'Replace the old collection with the new one
    'Set MyItems = temp
    
    For Each item In temp
        Debug.Print item.config & vbTab & item.Lable & vbTab; item.Qty & vbTab & item.Description & vbTab & item.PartNum
    Next
    
    'Get Tree BOM
    
    Dim treeBom()   As BomPosition
    GetTreeBom temp, treeBom
    
    'is SLDPRT attached Direct to SLDASM (Tong)
    IsDirectSLDASM treeBom, bom
    
    
    'Phan loai BOM VTGC, SPTC
    Dim bomVTTC() As BomPosition
    Dim bomGC() As BomPosition
    getBOMSplit bom, bomVTTC, bomGC
    
    'write BOM to Excel.Application
    Set exApp = GetObject("", "Excel.Application")
    Dim closeWorkbook As Boolean
    closeWorkbook = IsWorkbookOpen(exApp, FILE_PATH)
    Set exWorkbook = exApp.Workbooks.Open(FILE_PATH)
    '
    Dim RetTree     As String
    RetTree = SaveBOMInExcelWithThumbNail(treeBom, "BOM TreeView")
    
    Dim RetGC       As String
    RetGC = SaveBOMInExcelWithThumbNail(bomGC, "BOM GC")
    
    Dim RetVTTC     As String
    RetVTTC = SaveBOMInExcelWithThumbNail(bomVTTC, "BOM VTTC")
    
    ' save          as file
    Dim newFilePath As String
    Dim fileVersion As Integer
    Dim filePathSave As String
    filePathSave = THUMNAIL_PATH & Folder & "\"
    fileVersion = 1
    newFilePath = filePathSave & Header & "_" & fileVersion & ".xlsx"
    
    ' Check if the file exists
    While Len(Dir(newFilePath)) <> 0
        ' If the file exists, increment the version number and create a new file path
        fileVersion = fileVersion + 1
        newFilePath = filePathSave & Header & "_" & fileVersion & ".xlsx"
    Wend
    
    exWorkbook.SaveAs newFilePath
    
    swApp.SendMsgToUser "BOM save at : " & newFilePath
  '
Else
    MsgBox "Please open assembly"
End If

End Sub

Sub SetCompVisib(swComp As SldWorks.Component2, level As Long, prefix As String, parent As String)
    
    Dim vChildArray As Variant
    Dim swChildComp As SldWorks.Component2
    Dim MyItem      As New Class1
    Dim part        As String
    Dim filePath    As String
    filePath = swComp.GetPathName()
    Dim PartNameFromPath As String
    PartNameFromPath = Mid(filePath, InStrRev(filePath, "\") + 1, InStrRev(filePath, ".") - InStrRev(filePath, "\") - 1)
    
    If swComp.GetSuppression() <> swComponentSuppressionState_e.swComponentSuppressed And Not swComp.ExcludeFromBOM Then
        MyItem.PathFile = filePath
        part = GetProperties(swComp, swComp.ReferencedConfiguration, "Number", PartNameFromPath)
        MyItem.PartNum = part        'Mid(filePath, InStrRev(filePath, "\") + 1, InStrRev(filePath, ".") - InStrRev(filePath, "\") - 1)
        MyItem.Material = GetProperties(swComp, swComp.ReferencedConfiguration, "Material", PartNameFromPath)
        MyItem.HeatTreatment = GetProperties(swComp, swComp.ReferencedConfiguration, "Heat Treatment", PartNameFromPath)
        MyItem.SurfaceProtection = GetProperties(swComp, swComp.ReferencedConfiguration, "Surface Protection", PartNameFromPath)
        MyItem.SurfaceFinish = GetProperties(swComp, swComp.ReferencedConfiguration, "Surface Finish", PartNameFromPath)
        MyItem.Comment = GetProperties(swComp, swComp.ReferencedConfiguration, "Comment", PartNameFromPath)
        MyItem.Weight = GetProperties(swComp, swComp.ReferencedConfiguration, "Weight", PartNameFromPath)
        MyItem.Length = GetProperties(swComp, swComp.ReferencedConfiguration, "Length", PartNameFromPath)
        MyItem.drawingNo = GetProperties(swComp, swComp.ReferencedConfiguration, "Drawing No", PartNameFromPath)
        MyItem.config = swComp.ReferencedConfiguration
        'add parent
        MyItem.myParent = parent
        
        For Each item In MyItems
            If item.PartNum = MyItem.PartNum And item.myParent = MyItem.myParent Then
                item.Qty = item.Qty + 1
                Exit Sub
            End If
        Next
        
        Dim cntMyItems  As Long
        cntMyItems = MyItems.Count
        
        Dim lbPre       As String
        If cntMyItems > 1 Then
            lbPre = MyItems(cntMyItems).Lable
            Dim lbLastVar As Variant
            lbLastVar = SplitString(lbPre, ".")
            Dim lbCurVar As Variant
            lbCurVar = SplitString(prefix, ".")
            If UBound(lbLastVar) = UBound(lbCurVar) Then
                If lbCurVar(UBound(lbCurVar)) - lbLastVar(UBound(lbLastVar)) >= 2 Then
                    lbCurVar(UBound(lbCurVar)) = lbLastVar(UBound(lbLastVar)) + 1
                    prefix = CStr(lbCurVar(1))
                    For l = 2 To UBound(lbCurVar)
                        prefix = prefix & "." & CStr(lbCurVar(l))
                    Next l
                End If
            ElseIf UBound(lbLastVar) > UBound(lbCurVar) Then
                If lbCurVar(UBound(lbCurVar)) - lbLastVar(UBound(lbCurVar)) >= 2 Then
                    lbCurVar(UBound(lbCurVar)) = lbLastVar(UBound(lbCurVar)) + 1
                    prefix = CStr(lbCurVar(1))
                    For l = 2 To UBound(lbCurVar)
                        prefix = prefix & "." & CStr(lbCurVar(l))
                    Next l
                End If
            End If
        End If
        
        MyItem.Lable = prefix
        MyItem.Qty = 1
        MyItem.Description = GetProperties(swComp, swComp.ReferencedConfiguration, "Description", PartNameFromPath)
        If MyItem.PartNum Like "00*" Then
            MyItem.category = "SPTC (San Pham Tieu Chuan)"
        Else
            MyItem.category = "VTGC (Vat Tu Gia Cong)"
        End If
        MyItems.Add MyItem
        swComp.Visible = swComponentVisible
        vChildArray = swComp.GetChildren
        Dim i           As Long
        For i = 0 To UBound(vChildArray)
            
            Dim j       As Long
            j = i + 1
            
            Dim label   As String
            
            If prefix = "" Then
                label = CStr(j)
            Else
                label = prefix & "." & CStr(j)
            End If
            Set swChildComp = vChildArray(i)
            SetCompVisib swChildComp, j, label, GetProperties(swComp, swComp.ReferencedConfiguration, "Number", PartNameFromPath)
        Next i
    End If
End Sub

Function GetTreeBom(MyItems As Collection, ByRef treeBom() As BomPosition)
    
    For Each MyItem In MyItems
        If (Not treeBom) = -1 Then
            ReDim treeBom(0)
        Else
            ReDim Preserve treeBom(UBound(treeBom) + 1)
        End If
        Dim cntTreebom As Long
        cntTreebom = UBound(treeBom)
        treeBom(cntTreebom).ModelPath = MyItem.PathFile
        treeBom(cntTreebom).indent = MyItem.Lable
        treeBom(cntTreebom).Description = MyItem.Description
        treeBom(cntTreebom).drawingNo = MyItem.drawingNo
        treeBom(cntTreebom).category = MyItem.category
        treeBom(cntTreebom).partNumber = MyItem.PartNum
        treeBom(cntTreebom).Material = MyItem.Material
        treeBom(cntTreebom).HeatTreatment = MyItem.HeatTreatment
        treeBom(cntTreebom).SurfaceProtection = MyItem.SurfaceProtection
        treeBom(cntTreebom).SurfaceFinish = MyItem.SurfaceFinish
        treeBom(cntTreebom).Comment = MyItem.Comment
        treeBom(cntTreebom).Weight = MyItem.Weight
        treeBom(cntTreebom).Length = MyItem.Length
        treeBom(cntTreebom).Quantity = MyItem.Qty
        treeBom(cntTreebom).isWelment = MyItem.isWelment
        treeBom(cntTreebom).myParent = MyItem.myParent
    Next
    
End Function

Function GetFlatBom(assy As SldWorks.AssemblyDoc) As BomPosition()
    
    Dim bom()       As BomPosition
    
    Dim vComps      As Variant
    vComps = assy.GetComponents(False)
    
    Dim i           As Integer
    
    For i = 0 To UBound(vComps)
        
        Dim swComp  As SldWorks.Component2
        Set swComp = vComps(i)
        
        If swComp.GetSuppression() <> swComponentSuppressionState_e.swComponentSuppressed And Not swComp.ExcludeFromBOM Then
            
            Dim bomPos As Integer
            bomPos = FindBomPosition(bom, swComp)
            
            If bomPos = -1 Then
                
                If (Not bom) = -1 Then
                    ReDim bom(0)
                Else
                    ReDim Preserve bom(UBound(bom) + 1)
                End If
                
                bomPos = UBound(bom)
                
                bom(bomPos).ModelPath = swComp.GetPathName()
                bom(bomPos).Configuration = swComp.ReferencedConfiguration
                bom(bomPos).Quantity = 1
                
                Dim PartNameFromPath As String
                PartNameFromPath = Mid(bom(bomPos).ModelPath, InStrRev(bom(bomPos).ModelPath, "\") + 1, InStrRev(bom(bomPos).ModelPath, ".") - InStrRev(bom(bomPos).ModelPath, "\") - 1)
                
                'son
                If PartNameFromPath Like "QP*" Then
                    If Not swComp.ReferencedConfiguration Like "Default" Then
                        If Not swComp.ReferencedConfiguration Like "Default<As Machined>" Then
                            swApp.SendMsgToUser PartNameFromPath & " Not active config TAB <Default>"
                            End
                        End If
                
                    End If
                End If
                'get properties
                Dim swCompModel As SldWorks.ModelDoc2
                Set swCompModel = swComp.GetModelDoc2()
                bom(bomPos).Description = GetPropertyValue(swCompModel, swComp.ReferencedConfiguration, "Description", PartNameFromPath)
                bom(bomPos).drawingNo = GetPropertyValue(swCompModel, swComp.ReferencedConfiguration, "Drawing No", PartNameFromPath)
                Dim part As String
                part = GetPropertyValue(swCompModel, swComp.ReferencedConfiguration, "Number", PartNameFromPath)
                bom(bomPos).partNumber = part
                bom(bomPos).Material = GetPropertyValue(swCompModel, swComp.ReferencedConfiguration, "Material", PartNameFromPath)
                bom(bomPos).HeatTreatment = GetPropertyValue(swCompModel, swComp.ReferencedConfiguration, "Heat Treatment", PartNameFromPath)
                bom(bomPos).SurfaceProtection = GetPropertyValue(swCompModel, swComp.ReferencedConfiguration, "Surface Protection", PartNameFromPath)
                bom(bomPos).SurfaceFinish = GetPropertyValue(swCompModel, swComp.ReferencedConfiguration, "Surface Finish", PartNameFromPath)
                bom(bomPos).Comment = GetPropertyValue(swCompModel, swComp.ReferencedConfiguration, "Comment", PartNameFromPath)
                bom(bomPos).Weight = GetPropertyValue(swCompModel, swComp.ReferencedConfiguration, "Weight", PartNameFromPath)
                bom(bomPos).Length = GetPropertyValue(swCompModel, swComp.ReferencedConfiguration, "Length", PartNameFromPath)
                '
                If swComp.GetPathName() Like "*.SLDASM" Then
                    bom(bomPos).isSLDASM = True
                Else
                    bom(bomPos).isSLDASM = False
                End If
                'add cutList
                If swComp.GetPathName() Like "*.SLDPRT" And part <> "" Then
                    ProcessCutLists swCompModel, swComp, bom, part, bomPos
                End If
                
            Else
                bom(bomPos).Quantity = bom(bomPos).Quantity + 1
            End If
            
        End If
        
    Next
    
    GetFlatBom = bom
    
End Function

Function FindBomPosition(bom() As BomPosition, comp As SldWorks.Component2) As Integer
    
    FindBomPosition = -1
    
    If (Not bom) <> -1 Then
        Dim i       As Integer
        
        For i = 0 To UBound(bom)
            If LCase(bom(i).ModelPath) = LCase(comp.GetPathName()) And LCase(bom(i).Configuration) = LCase(comp.ReferencedConfiguration) Then
                FindBomPosition = i
                Exit Function
            End If
        Next
    End If
    
End Function

Function GetProperties(comp As SldWorks.Component2, conf As String, prName As String, PartNameFromPath As String) As String
    
    On Error GoTo err_
    
    Dim refModel    As SldWorks.ModelDoc2
    Set refModel = comp.GetModelDoc2()
    
    'Dim filePath    As String
    'filePath = comp.GetPathName()
    'Dim PartNameFromPath As String
    'PartNameFromPath = Mid(filePath, InStrRev(filePath, "\") + 1, InStrRev(filePath, ".") - InStrRev(filePath, "\") - 1)
    
    GetProperties = GetPropertyValue(refModel, conf, prName, PartNameFromPath)
    Exit Function
    
err_:
    Debug.Print "Failed To extract quantity of " & comp.Name2 & ": " & Err.Description
    GetProperties = ""
    
End Function

'conf As String hard
Function GetPropertyValue(model As SldWorks.ModelDoc2, conf As String, prpName As String, PartNameFromPath As String) As String
    
    Dim confSpecPrpMgr As SldWorks.CustomPropertyManager
    Dim genPrpMgr   As SldWorks.CustomPropertyManager
    Dim prpVal  As String
    Dim prpResVal As String

    If model Is Nothing Then
        GetPropertyValue = ""
    'hard
    Else
        
        If (PartNameFromPath Like "00.*") Then
            Set confSpecPrpMgr = model.Extension.CustomPropertyManager(conf)
            If confSpecPrpMgr Is Nothing Then
                GetPropertyValue = ""
                Exit Function
            End If
            
            Set genPrpMgr = model.Extension.CustomPropertyManager("")
            confSpecPrpMgr.Get3 prpName, False, prpVal, prpResVal
            
            If prpResVal = "" Then
                genPrpMgr.Get3 prpName, False, prpVal, prpResVal
            End If
            GetPropertyValue = prpResVal
        
        ElseIf (conf = "Default" Or conf = "Default<As Machined>") Then
            Set confSpecPrpMgr = model.Extension.CustomPropertyManager(conf)

            If confSpecPrpMgr Is Nothing Then
                GetPropertyValue = ""
                Exit Function
            End If
            
            Set genPrpMgr = model.Extension.CustomPropertyManager("")
            confSpecPrpMgr.Get3 prpName, False, prpVal, prpResVal
            
            If prpResVal = "" Then
                genPrpMgr.Get3 prpName, False, prpVal, prpResVal
            End If
            GetPropertyValue = prpResVal
        Else

            Set confSpecPrpMgr = model.Extension.CustomPropertyManager("Default")

            If confSpecPrpMgr Is Nothing Then
                Set confSpecPrpMgr = model.Extension.CustomPropertyManager("Default<As Machined>")
                If confSpecPrpMgr Is Nothing Then
                    GetPropertyValue = ""
                    Exit Function
                End If
            End If
            
            Set genPrpMgr = model.Extension.CustomPropertyManager("")
            confSpecPrpMgr.Get3 prpName, False, prpVal, prpResVal
            
            If prpResVal = "" Then
                genPrpMgr.Get3 prpName, False, prpVal, prpResVal
            End If
            GetPropertyValue = prpResVal
            
        End If
    End If
    
End Function

Public Function SaveBOMInExcelWithThumbNail(bom() As BomPosition, sheet As String) As String
    
    Set exWorkSheet = exWorkbook.Sheets(sheet)
    
    If exApp Is Nothing Then
        SaveBOMInExcelWithThumbNail = "Unable To initialize the Excel application"
        Exit Function
    End If
    exApp.Visible = True
    
    If exWorkSheet Is Nothing Then
        SaveBOMInExcelWithThumbNail = "Unable To Get the active sheet"
        Exit Function
    End If
    
    If UBound(bom) = 0 Then
        SaveBOMInExcelWithThumbNail = "BOM has no rows!"
    End If
    
    'set column width
    exWorkSheet.Columns(1).ColumnWidth = Width
    
    'header
    If sheet = "BOM GC" Then
        exWorkSheet.cells(1, 7).Value = Header
    ElseIf sheet = "BOM TreeView" Then
        exWorkSheet.cells(1, 7).Value = Header
    Else
        exWorkSheet.cells(1, 8).Value = Header
    End If
    
Skipper:
    For i = 0 To UBound(bom)
        
        'insert image
        Dim partNumber As String
        partNumber = bom(i).partNumber
        
        Dim imagePath As String
        If bom(i).isWelment And bom(i).Description <> "" Then
            Dim variantPart As Variant
            variantPart = SplitString(bom(i).Description, "-")
            If Not IsNull(variantPart) And UBound(variantPart) >= 2 Then
                imagePath = THUMNAIL_PATH_WELMENT & variantPart(1) & "\" & variantPart(2) & "\" & variantPart(1) & "-" & variantPart(2) & ".jpg"
            Else
                imagePath = THUMNAIL_PATH_WELMENT & "abc.jpg"
            End If
            
        Else
            Dim fileExtension As String
            fileExtension = Right(bom(i).ModelPath, Len(bom(i).ModelPath) - InStrRev(bom(i).ModelPath, "."))
            Dim img As String
            img = FindFilesWithPrefix(Left(partNumber, 15), thumbnailPath)
            If IsNull(img) Or Len(img) = 0 Then
                imagePath = thumbnailPath & "abc.jpg"
            Else
                imagePath = thumbnailPath & img
                'thumbnailPath & partNumber & "." & fileExtension & ".jpg"
            End If
            
        End If
        exWorkSheet.Rows(i + RowStart).RowHeight = Height
        'end insert image
        If bom(i).isWelment And sheet <> "BOM TreeView" Then
            Dim p   As String
            p = bom(i).PartNumParentWelment
            Dim QtyParent As Long
            QtyParent = GetQtyParent(p, bom)
            Dim CurrentQty As Long
            CurrentQty = bom(i).Quantity
            If QtyParent > 1 Then
                bom(i).Quantity = CLng(CurrentQty) * CLng(QtyParent)
            End If
        End If
        'insert data to template
        If sheet = "BOM GC" Then
            InsertPictureInRange exWorkSheet, imagePath, exWorkSheet.Range("E" & i + RowStart & ":E" & i + RowStart)
            exWorkSheet.cells(i + RowStart, 1).Value = i + 1
            exWorkSheet.cells(i + RowStart, 2).Value = bom(i).category
            exWorkSheet.cells(i + RowStart, 3).Value = bom(i).partNumber
            exWorkSheet.cells(i + RowStart, 4).Value = bom(i).drawingNo
            exWorkSheet.cells(i + RowStart, 6).Value = bom(i).Description
            exWorkSheet.cells(i + RowStart, 7).Value = bom(i).Material
            exWorkSheet.cells(i + RowStart, 8).Value = bom(i).Quantity
            exWorkSheet.cells(i + RowStart, 9).Value = bom(i).HeatTreatment
            exWorkSheet.cells(i + RowStart, 10).Value = bom(i).SurfaceProtection
            exWorkSheet.cells(i + RowStart, 11).Value = bom(i).SurfaceFinish
            exWorkSheet.cells(i + RowStart, 12).Value = bom(i).Comment
            exWorkSheet.cells(i + RowStart, 13).Value = bom(i).Length
            exWorkSheet.cells(i + RowStart, 14).Value = bom(i).Weight
            Dim gc As Integer
            For gc = 1 To 14
                Dim columNamegc As String
               columNamegc = exWorkSheet.cells(i + RowStart, gc).Address(False, False)
               exWorkSheet.Range(columNamegc & ":" & columNamegc).BorderAround _
    ColorIndex:=-4105, Weight:=2
            
            Next gc
            
        ElseIf sheet = "BOM TreeView" Then
            InsertPictureInRange exWorkSheet, imagePath, exWorkSheet.Range("E" & i + RowStart & ":E" & i + RowStart)
            'format to text
            exWorkSheet.cells(i + RowStart, 1).NumberFormat = "@"
            exWorkSheet.cells(i + RowStart, 1).Value = bom(i).indent
            exWorkSheet.cells(i + RowStart, 2).Value = bom(i).category
            exWorkSheet.cells(i + RowStart, 3).Value = bom(i).partNumber
            exWorkSheet.cells(i + RowStart, 4).Value = bom(i).drawingNo
            exWorkSheet.cells(i + RowStart, 6).Value = bom(i).Description
            exWorkSheet.cells(i + RowStart, 7).Value = bom(i).Material
            exWorkSheet.cells(i + RowStart, 8).Value = bom(i).Quantity
            exWorkSheet.cells(i + RowStart, 9).Value = bom(i).HeatTreatment
            exWorkSheet.cells(i + RowStart, 10).Value = bom(i).SurfaceProtection
            exWorkSheet.cells(i + RowStart, 11).Value = bom(i).SurfaceFinish
            exWorkSheet.cells(i + RowStart, 12).Value = bom(i).Comment
            exWorkSheet.cells(i + RowStart, 13).Value = bom(i).Length
            exWorkSheet.cells(i + RowStart, 14).Value = bom(i).Weight
            Dim tr As Integer
            For tr = 1 To 14
                Dim columNametr As String
               columNametr = exWorkSheet.cells(i + RowStart, tr).Address(False, False)
               exWorkSheet.Range(columNametr & ":" & columNametr).BorderAround _
    ColorIndex:=-4105, Weight:=2
            
            Next tr
            
            
        Else
            InsertPictureInRange exWorkSheet, imagePath, exWorkSheet.Range("F" & i + RowStart & ":F" & i + RowStart)
            exWorkSheet.cells(i + RowStart, 1).Value = i + 1
            exWorkSheet.cells(i + RowStart, 2).Value = bom(i).category
            exWorkSheet.cells(i + RowStart, 3).Value = bom(i).chunggLoai
            exWorkSheet.cells(i + RowStart, 4).Value = bom(i).partNumber
            exWorkSheet.cells(i + RowStart, 5).Value = bom(i).drawingNo
            exWorkSheet.cells(i + RowStart, 7).Value = bom(i).Description
            exWorkSheet.cells(i + RowStart, 8).Value = bom(i).Material
            exWorkSheet.cells(i + RowStart, 9).Value = bom(i).Quantity
            exWorkSheet.cells(i + RowStart, 10).Value = bom(i).HeatTreatment
            exWorkSheet.cells(i + RowStart, 11).Value = bom(i).SurfaceProtection
            exWorkSheet.cells(i + RowStart, 12).Value = bom(i).SurfaceFinish
            exWorkSheet.cells(i + RowStart, 13).Value = bom(i).Comment
            exWorkSheet.cells(i + RowStart, 14).Value = bom(i).Length
            exWorkSheet.cells(i + RowStart, 15).Value = bom(i).Weight
            Dim y As Integer
            For y = 1 To 15
                Dim columName As String
               columName = exWorkSheet.cells(i + RowStart, y).Address(False, False)
               exWorkSheet.Range(columName & ":" & columName).BorderAround _
    ColorIndex:=-4105, Weight:=2
            
            Next y
        End If
    Next i
    
End Function

Sub InsertPictureInRange(ActiveSheet As Object, PictureFileName As String, TargetCells As Object)
    On Error GoTo err_
    ' inserts a picture and resizes it to fit the TargetCells range
    Dim p           As Object, t As Double, l As Double, w As Double, h As Double
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If Dir(PictureFileName) = "" Then Exit Sub
    ' import picture
    'Set p = ActiveSheet.Pictures.Insert(PictureFileName)
    
    With TargetCells
        t = .Top
        l = .Left
        w = .Offset(0, .Columns.Count).Left - .Left
        h = .Offset(.Rows.Count, 0).Top - .Top
    End With
    
    With p
    ActiveSheet.Shapes.AddPicture _
                                  FileName:=PictureFileName, _
                                  LinkToFile:=False, _
                                  SaveWithDocument:=True, _
                                  Left:=l + 3, _
                                  Top:=t + 3, _
                                  Width:=w - 5, _
                                  Height:=h - 5
    
    End With
    
    Set TargetCells = Nothing
    Set p = Nothing
    ' determine positions
    'With TargetCells
    '    t = .Top
    '    l = .Left
    '    w = .Offset(0, .Columns.Count).Left - .Left
    '    h = .Offset(.Rows.Count, 0).Top - .Top
    'End With
    ' position picture
    'With p
    '   .Top = t
    '   .Left = l
    '   .Width = w
    '   .Height = h
    'End With
    'Set p = Nothing
err_:
    Exit Sub
End Sub

Function IsWorkbookOpen(xlApp As Object, filePath As String) As Boolean
    
    Dim i           As Integer
    
    For i = 1 To xlApp.Workbooks.Count
        If LCase(xlApp.Workbooks(i).FullName) = LCase(filePath) Then
            IsWorkbookOpen = True
            Exit Function
        End If
    Next
    
    IsWorkbookOpen = False
    
End Function

Function GetQtyParent(part As String, bom() As BomPosition) As Long
    Dim i           As Long
    For i = 0 To UBound(bom)
        If bom(i).partNumber = part Then
            GetQtyParent = bom(i).Quantity
            Exit Function
        Else
            GetQtyParent = 1
        End If
        
    Next i
    
End Function

Function getBOMSplit(bom() As BomPosition, ByRef bomVTTC() As BomPosition, ByRef bomGC() As BomPosition)
    
    For i = 0 To UBound(bom)
        Dim cntVTTC As Integer
        Dim part    As String
        part = bom(i).partNumber
        'end
        
        '  And (bom(i).isParentOfWelment = True Or bom(i).isSLDASM = true
        
        If (part = "" And bom(i).isWelment = True) Or part <> "" Then
            If (part Like "00.*" And (bom(i).isSLDASM = True Or bom(i).isDirect = True)) Then
                
                If (Not bomVTTC) = -1 Then
                    ReDim bomVTTC(0)
                Else
                    ReDim Preserve bomVTTC(UBound(bomVTTC) + 1)
                End If
                cntVTTC = UBound(bomVTTC)
                
                bomVTTC(cntVTTC) = bom(i)
                bomVTTC(cntVTTC).category = "SPTC (San Pham Tieu Chuan)"
                
                If part Like "00.01.*" Then
                    bomVTTC(cntVTTC).chunggLoai = "Dien"
                    
                ElseIf part Like "00.02.*" Or part Like "00.00.*" Then
                    bomVTTC(cntVTTC).chunggLoai = "Co khi"
                    
                ElseIf part Like "00.03.*" Then
                    bomVTTC(cntVTTC).chunggLoai = "Thuy Luc"
                    
                ElseIf part Like "00.04.*" Then
                    bomVTTC(cntVTTC).chunggLoai = "Khi Nen"
                Else
                    bomVTTC(cntVTTC).chunggLoai = part
                End If

            ElseIf Not part Like "00.*" And bom(i).isParentOfWelment = False And bom(i).isSLDASM = False Then
                If (Not bomGC) = -1 Then
                    ReDim bomGC(0)
                Else
                    ReDim Preserve bomGC(UBound(bomGC) + 1)
                End If
                cntGC = UBound(bomGC)
                
                bomGC(cntGC) = bom(i)
                bomGC(cntGC).category = "VTGC (Vat Tu Gia Cong)"
            Else
                
            End If
        End If
    Next i
End Function

Function ProcessCutLists(model As SldWorks.ModelDoc2, swComp As SldWorks.Component2, ByRef bomwel() As BomPosition, PartNum As String, idexParent As Integer)
    
    Dim swFeat      As SldWorks.Feature
    
    Set swFeat = model.FirstFeature
    
    Dim idx         As Integer
    idx = 0
    
    Do While Not swFeat Is Nothing
        
        Dim swBodyFolder As SldWorks.BodyFolder
        
        If swFeat.GetTypeName2() = "CutListFolder" Then
            Set swBodyFolder = swFeat.GetSpecificFeature2
            
            idx = idx + 1
            
            'Qty
            Dim bodyCount As Long
            bodyCount = swBodyFolder.GetBodyCount
            
            If bodyCount <> 0 Then
                
                Dim CustomManager As SldWorks.CustomPropertyManager
                Set CustomManager = swFeat.CustomPropertyManager
                'Description
                Dim Desc As String
                Desc = GetPropertiesByName(CustomManager, False, " ", "Description")
                
                bomwel(idexParent).isParentOfWelment = False
                If Desc <> "Sheet" Then
                    'set isParentOfWelment = True for .SLDPRT
                    bomwel(idexParent).isParentOfWelment = True
                    
                    'add data to BOM
                    Dim cntBomwel As Integer
                    cntBomwel = UBound(bomwel)
                    
                    If (Not bomwel) = -1 Then
                        ReDim bomwel(0)
                    Else
                        ReDim Preserve bomwel(0 To cntBomwel + 1)
                    End If
                    
                    cntBomwel = UBound(bomwel)
                    
                    bomwel(cntBomwel).Quantity = bodyCount
                    bomwel(cntBomwel).Description = Desc
                    Dim PartCutList As String
                    PartCutList = GetPropertiesByName(CustomManager, False, " ", "Number")
                    
                    bomwel(cntBomwel).partNumber = PartCutList
                    bomwel(cntBomwel).isWelment = True
                    bomwel(cntBomwel).PartNumParentWelment = PartNum
                    'Lengh
                    Dim Length As String
                    Length = GetPropertiesByName(CustomManager, False, " ", "LENGTH")
                    bomwel(cntBomwel).Length = Length
                    
                    'Material
                    Dim Material As String
                    Material = GetPropertiesByName(CustomManager, False, " ", "MATERIAL")
                    bomwel(cntBomwel).Material = Material
                    
                    'DrawingNo
                    Dim drawingNo As String
                    drawingNo = GetPropertiesByName(CustomManager, False, " ", "Drawing No")
                    bomwel(cntBomwel).drawingNo = drawingNo
                End If
                
            End If
            
        End If
        
        Set swFeat = swFeat.GetNextFeature
        
    Loop
    
End Function

Function GetPropertiesByName(custPrpMgr As SldWorks.CustomPropertyManager, cached As Boolean, indent As String, prpName As String) As String
    
    Dim textexp     As String
    Dim evalval     As String
    
    custPrpMgr.Get2 prpName, textexp, evalval
    GetPropertiesByName = evalval
End Function

Function SplitString(text As String, delimiter As String) As Variant
    
    Dim vSplit      As Variant
    ReDim vSplit(0)        ' Initialize empty array
    
    ' Handle empty string or delimiter
    If text = "" Or delimiter = "" Then
        Exit Function
    End If
    
    Dim start       As Long, i As Long
    start = 1
    
    For i = 1 To Len(text)
        If Mid(text, i, 1) = delimiter Then
            ReDim Preserve vSplit(UBound(vSplit) + 1)
            vSplit(UBound(vSplit)) = Mid(text, start, i - start)
            start = i + 1
        End If
    Next i
    
    ' Add the last substring if delimiter not at the end
    If start <= Len(text) Then
        ReDim Preserve vSplit(UBound(vSplit) + 1)
        vSplit(UBound(vSplit)) = Mid(text, start)
    End If
    
    SplitString = vSplit
    
End Function

Function IsDirectSLDASM(treeBom() As BomPosition, bom() As BomPosition)
    Dim i As Long
    For i = 0 To UBound(bom)
        Dim j As Long
        For j = 0 To UBound(treeBom)
            If bom(i).partNumber = treeBom(j).partNumber And bom(i).isSLDASM = False Then
                Dim variantLable As Variant
                variantLable = SplitString(treeBom(j).indent, ".")
                If (Not IsNull(variantLable) And UBound(variantLable) = 2) Or treeBom(j).myParent Like "QP.*" Then
                    bom(i).isDirect = True
                Else
                    bom(i).isDirect = False
                End If
                'Exit Function
            End If
        Next j
    Next i
End Function

Function FindFilesWithPrefix(filePrefix As String, folderPath As String) As String
    
    Dim file As String
    file = Dir(folderPath & "*.jpg")
    Do While file <> ""
        If Left(file, 15) = filePrefix Then
            FindFilesWithPrefix = file
            Exit Function
        End If
        file = Dir()
    Loop

End Function


Attribute VB_Name = "ExcelBuilderModule"
''' <summary>Module: ExcelBuilderModule</sammary>
''' <remarks>
''' Requires
''' - Microsoft ActiveX Data Objects X.X Library
''' - Microsoft VBScript Regular Expressions 5.5
''' - Microsoft Scripting Runtime
''' </remarks>


Option Explicit

Public Enum ValueType
    TYPE_INCREMENT = 1
    TYPE_STRING
    TYPE_LIST
End Enum

''' <summary>Normalize Sheet Name</sammary>
Function EscapeSheetName(ByRef sheetName As String) As String
    Dim reservedChars As String
    Dim i As Long
    Dim escapedName As String

    reservedChars = "\/?*" & Chr(34) & "<>" & "|"
    escapedName = sheetName

    For i = 1 To Len(reservedChars)
        escapedName = Replace(escapedName, Mid(reservedChars, i, 1), "_")
    Next i

    EscapeSheetName = Left(escapedName, 31)
End Function

''' <summary>Join all strings in Collection</sammary>
Function JoinCollection(contents As Collection) As String
    Dim item As Variant
    Dim str As String
    
    str = ""
    For Each item In contents
        str = str & item & vbLf
    Next
    
    JoinCollection = str
End Function

'''<summary>Convert Markdown to Sheet</summary>
'''<params id="filePath"></params>
Sub Convert(ByVal filePath As String, ByVal config As ColumnConfig)
    Dim adoStream As New ADODB.Stream
    Dim rowNumber As Long
    Dim line As String
    Dim mc As MatchCollection
    Dim titleRow As Long
    
    rowNumber = 1

    With adoStream
        .Open
        .Type = adTypeText
        .Charset = "UTF-8"
        .LineSeparator = adLF
        .LoadFromFile filePath  ' Require Open
        
        ' set name
        Do Until .EOS
            line = .ReadText(-2)
            
            With New RegExp
                .Pattern = "^#[^#]\s*(\S.*)\s*$"
                Set mc = .Execute(line)
                If mc.Count > 0 Then
                    ActiveWorkbook.Sheets.Add
                    ActiveSheet.Name = EscapeSheetName(mc(0).SubMatches(0))
                    Exit Do
                End If
            End With
        Loop
        
        If .EOS Then
            MsgBox "Can not find title."
            Exit Sub
        End If
            
        ' set summary
        Dim hasSammary As Boolean
        hasSammary = False
        Do Until .EOS
            line = .ReadText(-2)
            
            With New RegExp
                .Pattern = "^##\s*Summary\s*$"
                If .Test(line) Then
                    hasSammary = True
                    Exit Do
                End If
            End With
        Loop
               
        If hasSammary Then
            With Cells(rowNumber, 1)
                .Value = "Summary"
                .Font.Bold = True
                rowNumber = rowNumber + 2
            End With
        End If
        
        ' read summary lines
        Do Until .EOS
            line = .ReadText(-2)
            
            With New RegExp
                .Pattern = "^##\s*(List|Steps|Rows)\s*$"
                If .Test(line) Then
                    rowNumber = rowNumber + 1
                    Exit Do
                End If
            End With
            
            If Len(Trim(line)) > 0 Then
                Cells(rowNumber, 1).Value = RTrim(line)
                rowNumber = rowNumber + 1
            End If
        Loop

        ' read steps
        Dim steps As New Collection
        Dim stepDict As New Scripting.Dictionary
        Dim cond As ColumnCondition
        Dim item As StepItem
        Dim itemType As ValueType
        Dim title As String
        Set item = Nothing
        
        itemType = TYPE_STRING
        Do While True
            line = .ReadText(-2)
            
            If .EOS Then
                If Not item Is Nothing Then
                    stepDict.Add item.title, item.GetContent()
                End If
                If stepDict.Count > 0 Then
                    steps.Add stepDict
                End If
                Exit Do
            End If
            
            With New RegExp
                .Pattern = "^\s*---\s*$"
                If .Test(line) Then
                    If Not item Is Nothing Then
                        stepDict.Add item.title, item.GetContent()
                        If itemType = TYPE_LIST Then
                            If config.conditions.item(title).maxCount < item.GetContent().Count Then
                                config.conditions.item(title).maxCount = item.GetContent().Count
                            End If
                        End If
                        Set item = Nothing
                    End If
                    If stepDict.Count > 0 Then
                        steps.Add stepDict
                        Set stepDict = New Scripting.Dictionary
                    End If
                    GoTo CONTINUE
                End If
                
                .Pattern = "^#{3,}\s*(\S.*\S|\S)\s*$"
                Set mc = .Execute(line)
                If mc.Count > 0 Then
                    If Not item Is Nothing Then
                        stepDict.Add item.title, item.GetContent()
                        If itemType = TYPE_LIST Then
                            If config.conditions.item(title).maxCount < item.GetContent().Count Then
                                config.conditions.item(title).maxCount = item.GetContent().Count
                            End If
                        End If
                    End If
                    
                    title = mc(0).SubMatches(0)
                    If config.conditions.Exists(title) Then
                        itemType = config.conditions.item(title).value_type
                    Else
                        itemType = TYPE_STRING
                        Set cond = New ColumnCondition
                        config.conditions.Add title, cond
                    End If
                    
                    Set item = New StepItem
                    item.Init title, itemType
                    GoTo CONTINUE
                End If
                
                If Not item Is Nothing Then
                    If Len(Trim(line)) > 0 Then
                        item.AddContentLine (RTrim(line))
                    End If
                End If
            End With
CONTINUE:
        Loop
        
        .Close
        
        If steps.Count = 0 Then
            Exit Sub
        End If
        
        ' column header
        Dim allCond As Scripting.Dictionary
        Dim key As Variant
        Dim colNumber As Long
        Dim maxCol As Long
        Dim listIndex As Long
        colNumber = 1
        titleRow = rowNumber
        
        Set allCond = config.AllConditions()
        
        For Each key In allCond.Keys
            Set cond = config.AllConditions().item(key)
            With columns(colNumber)
                .ColumnWidth = cond.width
                .HorizontalAlignment = cond.alignment
            End With
            Select Case cond.value_type
            Case ValueType.TYPE_INCREMENT
                With Cells(rowNumber, colNumber)
                    .Value = key
                    .Font.Bold = True
                    .HorizontalAlignment = xlCenter
                End With
                colNumber = colNumber + 1
            Case ValueType.TYPE_LIST
                For listIndex = 1 To cond.maxCount
                    With columns(colNumber + listIndex)
                        .ColumnWidth = cond.width
                        .HorizontalAlignment = cond.alignment
                    End With
                    With Cells(rowNumber, colNumber + listIndex - 1)
                        .Value = key & listIndex
                        .Font.Bold = True
                        .HorizontalAlignment = xlCenter
                    End With
                Next
                colNumber = colNumber + cond.maxCount
            Case ValueType.TYPE_STRING
                With Cells(rowNumber, colNumber)
                    .Value = key
                    .Font.Bold = True
                    .HorizontalAlignment = xlCenter
                End With
                colNumber = colNumber + 1
            End Select
            
        Next
        
        maxCol = colNumber - 1
        
        rowNumber = rowNumber + 1
        
        ' table body
        Dim obj As Object
        Dim content As Collection
        For Each stepDict In steps
            colNumber = 1
            For Each key In config.prepend_columns
                If config.prepend_columns.item(key).value_type = TYPE_INCREMENT Then
                    Cells(rowNumber, colNumber).Value = config.prepend_columns.item(key).initial_value
                    config.prepend_columns.item(key).initial_value = config.prepend_columns.item(key).initial_value + 1
                End If
                colNumber = colNumber + 1
            Next
            For Each key In config.conditions.Keys
                Select Case config.conditions.item(key).value_type
                Case TYPE_LIST
                    If stepDict.Exists(key) Then
                        Set content = stepDict.item(key)
                        listIndex = 1
                        If content.Count > 0 Then
                            For listIndex = 1 To content.Count
                                With New RegExp
                                    .Pattern = "^-\s+"
                                    Cells(rowNumber, colNumber + listIndex - 1).Value = .Replace(content.item(listIndex), "")
                                End With
                            Next
                        End If
                    End If
                    colNumber = colNumber + config.conditions.item(key).maxCount
                Case TYPE_STRING
                    If stepDict.Exists(key) Then
                        Set content = stepDict.item(key)
                        Cells(rowNumber, colNumber).Value = JoinCollection(content)
                    End If
                    colNumber = colNumber + 1
                End Select
            Next
            rowNumber = rowNumber + 1
        Next
        
        ' format borders
        Dim index As Long
        With Range(Cells(titleRow, 1), Cells(rowNumber - 1, maxCol))
            For index = xlEdgeLeft To xlInsideHorizontal
                With .Borders(index)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
            Next
        End With
    End With
End Sub

'''<summary>Control Build Process</summary>
Sub Build(control As IRibbonControl)
    Dim filePath As Variant
    Dim configPath As String
    Dim config As New ColumnConfig
    
    filePath = Application.GetOpenFilename("markdown,*.md")
    
    If filePath = False Then
        Exit Sub
    End If
    
    configPath = Replace(filePath, Dir(filePath), "config\columns.xlsx")
    
    If Dir(configPath) <> "" Then
        config.Parse (configPath)
    End If
        
    Convert CStr(filePath), config
End Sub

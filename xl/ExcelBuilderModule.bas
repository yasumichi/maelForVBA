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
Sub Convert(filePath As String)
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
            
            Cells(rowNumber, 1).Value = RTrim(line)
            rowNumber = rowNumber + 1
        Loop

        ' read steps
        Dim steps As New Collection
        Dim stepDict As New Scripting.Dictionary
        Dim columns As New Scripting.Dictionary
        Dim item As StepItem
        Dim title As String
        Set item = Nothing
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
                    End If
                    
                    title = mc(0).SubMatches(0)
                    If Not columns.Exists(title) Then
                        columns.Add title, ""
                    End If
                    
                    Set item = New StepItem
                    item.Init title, TYPE_STRING
                    GoTo CONTINUE
                End If
                
                If Not item Is Nothing Then
                    item.AddContentLine (RTrim(line))
                End If
            End With
            'Cells(rowNumber, 1).Value = line
            'rowNumber = rowNumber + 1]
CONTINUE:
        Loop
        
        Dim key As Variant
        Dim colNumber As Long
        colNumber = 1
        For Each key In columns.Keys
            titleRow = rowNumber
            With Cells(rowNumber, colNumber)
                .Value = key
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
            End With
            colNumber = colNumber + 1
        Next
        
        rowNumber = rowNumber + 1
        
        Dim obj As Object
        For Each stepDict In steps
            colNumber = 1
            If stepDict.Count > 0 Then
                Dim index As Long
                Dim arr As Variant
                Dim content As Collection
                For index = 0 To stepDict.Count - 1
                    arr = stepDict.Items
                    Set content = arr(index)
                    Cells(rowNumber, colNumber).Value = JoinCollection(content)
                    colNumber = colNumber + 1
                Next
            End If
            rowNumber = rowNumber + 1
        Next
        
        ' format borders
        With Range(Cells(titleRow, 1), Cells(rowNumber - 1, columns.Count))
            For index = xlEdgeLeft To xlInsideHorizontal
                With .Borders(index)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
            Next
        End With
            
        .Close
    End With
End Sub

'''<summary>Control Build Process</summary>
Sub Build(control As IRibbonControl)
    Dim filePath As Variant
    
    filePath = Application.GetOpenFilename("markdown,*.md")
    
    If filePath = False Then
        Exit Sub
    End If
    
    Convert CStr(filePath)
End Sub

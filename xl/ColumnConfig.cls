VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' <summary>Class ColumnConfig</summary>
''' <remarks>
''' Requires
''' - Microsoft Scripting Runtime
''' </remarks>

Option Explicit

Public prepend_columns As Scripting.Dictionary
Public conditions As Scripting.Dictionary
Public append_columns As Scripting.Dictionary

''' <summary>default constructor</sammary>
Private Sub Class_Initialize()
    Set prepend_columns = New Scripting.Dictionary
    Set conditions = New Scripting.Dictionary
    Set append_columns = New Scripting.Dictionary
End Sub

''' <summary>list columns</sammary>
Public Function AllConditions() As Scripting.Dictionary
    Dim key As Variant
    Dim dict As New Scripting.Dictionary
    
    For Each key In prepend_columns.Keys
        dict.Add key, prepend_columns.item(key)
    Next
    
    For Each key In conditions.Keys
        dict.Add key, conditions.item(key)
    Next
    
    For Each key In append_columns.Keys
        dict.Add key, append_columns.item(key)
    Next

    Set AllConditions = dict
End Function


''' <summary>parse configuration of columns</sammary>
Public Sub Parse(configPath As String)
    Dim sh As Worksheet
    Dim coldic As Scripting.Dictionary
    Dim rowNumber As Long
    Dim colName As String
    Dim condition As ColumnCondition

    If Right(configPath, 5) <> ".xlsx" Then
        Debug.Print configPath & " is not supported."
        Exit Sub
    End If
    
    If Dir(configPath) = "" Then
        Debug.Print configPath & " is not exists."
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    With Workbooks.Open(configPath)
        With sh
        End With
        For Each sh In .Worksheets
            Set coldic = New Scripting.Dictionary
            With sh
                rowNumber = 2
                Do While .Cells(rowNumber, 1).Value <> ""
                    Set condition = New ColumnCondition
                    colName = .Cells(rowNumber, 1).Value
                    ' value_type
                    Select Case LCase(.Cells(rowNumber, 2).Value)
                    Case "increment"
                        condition.value_type = TYPE_INCREMENT
                        If .Cells(rowNumber, 3).Value <> "" Then
                            condition.initial_value = Val(.Cells(rowNumber, 3).Value)
                        End If
                    Case "list"
                        condition.value_type = TYPE_LIST
                    Case "string"
                        condition.value_type = TYPE_STRING
                    Case Else
                        condition.value_type = TYPE_STRING
                    End Select
                    
                    ' alignment
                    If .Cells(rowNumber, 4).Value <> "" Then
                        Select Case LCase(.Cells(rowNumber, 4).Value)
                        Case "left"
                            condition.alignment = xlLeft
                        Case "center"
                            condition.alignment = xlCenter
                        Case "right"
                            condition.alignment = xlRight
                        Case Else
                            condition.alignment = xlLeft
                        End Select
                    End If
                    
                    ' width
                    If .Cells(rowNumber, 5).Value <> "" Then
                        condition.width = Val(.Cells(rowNumber, 5).Value)
                    End If
                    
                    
                    coldic.Add colName, condition
                    rowNumber = rowNumber + 1
                    Set condition = Nothing
                Loop
                
            End With
            Select Case sh.Name
            Case "prepend"
                Set prepend_columns = coldic
            Case "column_conditions"
                Set conditions = coldic
            Case "append"
                Set append_columns = coldic
            End Select
            Set coldic = Nothing
        Next
        .Close
    End With
    
    Application.ScreenUpdating = True
End Sub


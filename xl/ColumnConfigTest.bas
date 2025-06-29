Attribute VB_Name = "ColumnConfigTest"
Option Explicit

Sub TestParse()
    Dim config As New ColumnConfig
    Dim filePath As String
    Dim key As Variant
    Dim cond As ColumnCondition
    
    filePath = Application.GetOpenFilename
    
    'Debug.Print Replace(filePath, Dir(filePath), "")
    
    config.parse (filePath)
    
    For Each key In config.prepend_columns.Keys
        Debug.Print CStr(key)
        Set cond = config.prepend_columns.item(key)
        Debug.Print "- " & cond.value_type
        Debug.Print "- " & cond.initial_value
        Debug.Print "- " & cond.alignment
        Debug.Print "- " & cond.width
    Next
    
    For Each key In config.conditions.Keys
        Debug.Print CStr(key)
        Set cond = config.conditions.item(key)
        Debug.Print "- " & cond.value_type
        Debug.Print "- " & cond.initial_value
        Debug.Print "- " & cond.alignment
        Debug.Print "- " & cond.width
    Next
    
    For Each key In config.append_columns.Keys
        Debug.Print CStr(key)
        Set cond = config.append_columns.item(key)
        Debug.Print "- " & cond.value_type
        Debug.Print "- " & cond.initial_value
        Debug.Print "- " & cond.alignment
        Debug.Print "- " & cond.width
    Next
End Sub

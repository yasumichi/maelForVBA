VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StepItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' <summary>Class StepItem</sammary>

Option Explicit

Public title As String
Private itemType As ValueType
Private contentLines As Collection
Private contentItems As Collection

''' <summary>Constructor</sammary>
Private Sub Class_Initialize()
    title = ""
    itemType = TYPE_STRING
    Set contentLines = New Collection
    Set contentItems = New Collection
End Sub

Public Sub Init(title_ As String, type_ As ValueType)
    title = title_
    itemType = type_
End Sub

''' <summary>Add line of Content</sammary>
''' <params id="content"></param>
Public Sub AddContentLine(content As String)
    If itemType = TYPE_STRING Then
        contentLines.Add (content)
    ElseIf itemType = TYPE_LIST Then
        contentItems.Add (content)
    End If
End Sub

''' <summary>Get lines of Content</sammary>
''' <returns>lines of Content</returns>
Public Function GetContent() As Collection
    If itemType = TYPE_STRING Then
        Set GetContent = contentLines
    ElseIf itemType = TYPE_LIST Then
        Set GetContent = contentItems
    End If
End Function

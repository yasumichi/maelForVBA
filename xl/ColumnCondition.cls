VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' <summary>Class ColumnCondition</summary>

Option Explicit

Public value_type As ValueType
Public initial_value As Long
Public width As Double
Public alignment As Long
Public maxCount As Long

''' <summary>Constructor</sammary>
Private Sub Class_Initialize()
    value_type = TYPE_STRING
    initial_value = 1
    width = 8.38
    alignment = xlGeneral
    maxCount = 1
End Sub

''' <summary>public initializer</sammary>
''' <params id="title_">column title</param>
''' <params id="type_">column value type</param>
Public Sub Init(vtype As ValueType, width_ As Double, alignment_ As Long)
    value_type = vtype
    width = width_
    alignment = alignment_
End Sub

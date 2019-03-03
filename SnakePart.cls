VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SnakePart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Properties
    row As Long
    column As Long
End Type

Private this As Properties

    Public Property Let row(ByVal value As Long)
        this.row = value
    End Property
    
    Public Property Get row() As Long
        row = this.row
    End Property
    
    Public Property Let column(ByVal value As Long)
        this.column = value
    End Property
    
    Public Property Get column() As Long
        column = this.column
    End Property
    
    Public Sub PropertiesSet(ByVal row As Long, ByVal column As Long)
        this.row = row
        this.column = column
    End Sub

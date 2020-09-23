VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCommandBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' cCommandBars is not a real object.  It is a proxy
' onto the mCommandBars.m_colCommandBar collection,
' given that the owning control exists.

Private m_hWnd As Long

Friend Sub fInit(ByVal hWnd As Long)
   m_hWnd = hWnd
End Sub

Public Property Get Count() As Long
Attribute Count.VB_Description = "Gets the number of CommandBars in the project."
Dim ctl As vbalCommandBar
   If (ControlFromhWnd(m_hWnd, ctl)) Then
      Count = ctl.BarCount()
   End If
End Property

Public Property Get Item(ByVal index As Variant) As cCommandBar
Attribute Item.VB_Description = "Gets the CommandBar with the specified index or key."
Attribute Item.VB_UserMemId = 0
Dim ctl As vbalCommandBar
   If (ControlFromhWnd(m_hWnd, ctl)) Then
      Set Item = ctl.BarItem(index)
   End If
End Property

Public Sub Remove(ByVal sKey As String)
Attribute Remove.VB_Description = "Removes the CommandBar with the specified index or key."
Dim ctl As vbalCommandBar
   If (ControlFromhWnd(m_hWnd, ctl)) Then
      ctl.BarRemove sKey
   End If
End Sub

Public Function Add(ByVal sKey As String, Optional ByVal sTitle As String = "") As cCommandBar
Attribute Add.VB_Description = "Adds a new CommandBar to the collection.  The mandatory Key must be globally unique across all CommandBars in your project."
Dim ctl As vbalCommandBar
   If (ControlFromhWnd(m_hWnd, ctl)) Then
      Set Add = ctl.BarAdd(sKey, sTitle)
   End If
End Function

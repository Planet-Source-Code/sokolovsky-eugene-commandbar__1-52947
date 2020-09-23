VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{EE7EB09A-2816-4406-A2B3-30431D62314F}#1.0#0"; "vbalCmdBar6.ocx"
Begin VB.Form frmMSMoneySample 
   Caption         =   "Money UI Demonstration"
   ClientHeight    =   6840
   ClientLeft      =   2640
   ClientTop       =   2445
   ClientWidth     =   11160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMSMoneySample.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   11160
   Begin VB.CommandButton cmdStyle 
      Caption         =   "Style"
      Height          =   435
      Left            =   180
      TabIndex        =   6
      Top             =   4635
      Width           =   1035
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "Font"
      Height          =   435
      Left            =   180
      TabIndex        =   5
      ToolTipText     =   "Demonstrates changing font"
      Top             =   4200
      Width           =   1035
   End
   Begin VB.PictureBox picSideBar 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   3
      Top             =   720
      Width           =   1515
   End
   Begin VB.PictureBox picRes 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   3300
      Picture         =   "frmMSMoneySample.frx":45A2
      ScaleHeight     =   825
      ScaleWidth      =   2400
      TabIndex        =   2
      Top             =   3420
      Visible         =   0   'False
      Width           =   2400
   End
   Begin vbalCmdBar6.vbalImageList ilsIcons 
      Left            =   1320
      Top             =   3600
      _ExtentX        =   953
      _ExtentY        =   953
      IconSizeX       =   24
      IconSizeY       =   24
      ColourDepth     =   24
      Size            =   29520
      Images          =   "frmMSMoneySample.frx":AD04
      Version         =   65536
      KeyCount        =   12
      Keys            =   "ABOUTÿPORTFOLIOÿMOREÿHOMEÿGOÿFORWARDÿBUDGETÿBILLSÿACCOUNTSÿVBACCELERATORÿBACKÿBASEEMPTY"
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   2835
      Left            =   1500
      TabIndex        =   1
      Top             =   840
      Width           =   6495
      ExtentX         =   11456
      ExtentY         =   5001
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.ComboBox cboAddress 
      Height          =   315
      Left            =   1980
      TabIndex        =   0
      Text            =   "http://vbaccelerator.com/"
      Top             =   420
      Width           =   2475
   End
   Begin vbalCmdBar6.vbalCommandBar cmdBar 
      Align           =   1  'Align Top
      Height          =   315
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   0
   End
   Begin vbalCmdBar6.vbalCommandBar cmdBar 
      Align           =   1  'Align Top
      Height          =   315
      Index           =   1
      Left            =   0
      Top             =   315
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   0
   End
   Begin vbalCmdBar6.vbalCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Index           =   2
      Left            =   0
      Top             =   6525
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   0
   End
   Begin vbalCmdBar6.vbalImageList ilsIcons16 
      Left            =   1980
      Top             =   3540
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   24
      Size            =   1148
      Images          =   "frmMSMoneySample.frx":12074
      Version         =   65536
      KeyCount        =   1
      Keys            =   ""
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMSMoneySample.frx":12510
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   1320
      Width           =   1200
   End
End
Attribute VB_Name = "frmMSMoneySample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Quit gently
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_CLOSE = &HF060&

' A generic recursive procedure to create command
' bars for all subitems of the specified key.
' This works if you set up your keys in the
' appropriate way.  It isn't particularly
' efficient, though.
Private Sub createCommandBarsFromKeys( _
      cmdBar As vbalCommandBar, _
      ByVal sStartKey As String, _
      btnOwner As cButton _
   )
Dim iBtn As Long
Dim bar As cCommandBar
Dim btn As cButton
Dim colStartKeyParts As Collection
Dim colParts As Collection

   Set colStartKeyParts = parseKey(sStartKey)

   With cmdBar.Buttons
      For iBtn = 1 To .Count
         Set btn = .Item(iBtn)
         If (InStr(btn.Key, sStartKey & ":") = 1) Then
            Set colParts = parseKey(btn.Key)
            If (colParts.Count = colStartKeyParts.Count + 1) Then
               If (bar Is Nothing) Then
                  Set bar = cmdBar.CommandBars.Add(sStartKey, sStartKey)
                  If Not (btnOwner Is Nothing) Then
                     btnOwner.bar = bar
                  End If
               End If
               bar.Buttons.Add btn
               ' recurse
               createCommandBarsFromKeys cmdBar, btn.Key, btn
            End If
         End If
      Next iBtn
   End With
   
End Sub

Private Function parseKey( _
      ByVal sKey As String _
   ) As Collection
Dim iPos As Long
Dim iNextPos As Long
Dim colParts As New Collection
      
   iPos = 1
   iNextPos = 1
   Do While (iNextPos > 0)
      iNextPos = InStr(iPos, sKey, ":")
      If (iNextPos > 0) Then
         colParts.Add Mid(sKey, iPos, iNextPos)
         iPos = iNextPos + 1
      End If
   Loop
   If (iPos > 0) Then
      colParts.Add Mid(sKey, iPos)
   End If
   
   Set parseKey = colParts
   
End Function

Private Sub createCommandBars()
   
   createCommandBarsFromKeys cmdBar(0), "MENU", Nothing
   
   createCommandBarsFromKeys cmdBar(0), "TOOLBAR", Nothing
   
   createCommandBarsFromKeys cmdBar(0), "STATUS", Nothing

End Sub

Private Sub createButtons()
Dim btn As cButton
Dim bar As cCommandBar
Dim btns As cCommandBarButtons
Dim i As Long

   With cmdBar(0)
      
      ' Add the buttons:
      With .Buttons
         
         ' Add top level menu buttons
         Set btn = .Add("MENU:FILE", , "&File")
         btn.ShowCaptionInToolbar = True
         
         Set btn = .Add("MENU:EDIT", , "&Edit")
         btn.ShowCaptionInToolbar = True
         
         Set btn = .Add("MENU:FAVOURITES", , "F&avourites")
         btn.ShowCaptionInToolbar = True
         
         Set btn = .Add("MENU:TOOLS", , "&Tools")
         btn.ShowCaptionInToolbar = True
         
         Set btn = .Add("MENU:ACCOUNTS", , "Accounts && &Bills")
         btn.ShowCaptionInToolbar = True
         
         Set btn = .Add("MENU:INVESTING", , "&Investing")
         btn.ShowCaptionInToolbar = True
         
         Set btn = .Add("MENU:PLANNER", , "&Planner")
         btn.ShowCaptionInToolbar = True
         
         Set btn = .Add("MENU:TAXES", , "Ta&xes")
         btn.ShowCaptionInToolbar = True
         
         Set btn = .Add("MENU:SHOPPING", , "&Shopping")
         btn.ShowCaptionInToolbar = True
         
         Set btn = .Add("MENU:HELP", , "&Help")
         btn.ShowCaptionInToolbar = True
         
         
         ' Add file menu buttons:
         .Add "MENU:FILE:NEW", , "&New"
         .Add "MENU:FILE:NEW:ACCOUNT", , "New &Account..."
         .Add "MENU:FILE:NEW:FILE", , "New &File..."
         .Add "MENU:FILE:OPEN", , "&Open", , , vbKeyO, vbCtrlMask
         .Add "MENU:FILE:CONVERT", , "Convert &Quicken File..."
         .Add "MENU:FILE:SEP1", , , eSeparator
         .Add "MENU:FILE:PASSWORD", , "Password &Manager"
         .Add "MENU:FILE:SEP2", , , eSeparator
         .Add "MENU:FILE:BACKUP", , "&Back Up..."
         .Add "MENU:FILE:RESTORE", , "&Restore Backup..."
         .Add "MENU:FILE:ARCHIVE", , "&Archive..."
         .Add "MENU:FILE:SEP3", , , eSeparator
         .Add "MENU:FILE:IMPORT", , "&Import..."
         .Add "MENU:FILE:EXPORT", , "&Export..."
         .Add "MENU:FILE:SEP4", , , eSeparator
         .Add "MENU:FILE:PRINTSETUP", , "P&rint Setup..."
         .Add "MENU:FILE:PRINT", , "&Print...", , , vbKeyP, vbCtrlMask
         .Add "MENU:FILE:PREVIEW", , "Print Pre&view"
         Set btn = .Add("MENU:FILE:MRUSEP", , , eSeparator)
         btn.Visible = False
         For i = 1 To 8
            Set btn = .Add("MENU:FILE:MRU" & i, , "Recent File")
            btn.Visible = False
         Next i
         .Add "MENU:FILE:SEP5", , , eSeparator
         .Add "MENU:FILE:EXIT", , "E&xit"
         
         ' Add Edit menu buttons:
         .Add "MENU:EDIT:UNDO", , "&Undo", , , vbKeyZ, vbCtrlMask
         .Add "MENU:EDIT:SEP1", , , eSeparator
         .Add "MENU:EDIT:CUT", , "Cu&t", , , vbKeyX, vbCtrlMask
         .Add "MENU:EDIT:COPY", , "&Copy", , , vbKeyC, vbCtrlMask
         .Add "MENU:EDIT:PASTE", , "&Paste", , , vbKeyV, vbCtrlMask
         
         ' Add Favourites:
         Set btn = .Add("MENU:FAVOURITES:1", , "(No Favourites Yet)")
         btn.Enabled = False
         
         ' Add Tools:
         .Add "MENU:TOOLS:FIND", , "&Find and Replace..."
         .Add "MENU:TOOLS:SEP1", , , eSeparator
         .Add "MENU:TOOLS:CALCULATOR", , "&Calculator..."
         .Add "MENU:TOOLS:DISCONNECT", , "&Disconnect"
         .Add "MENU:TOOLS:UPGRADE", , "&Upgrade..."
         .Add "MENU:TOOLS:CUSTOMISE", , "Customi&ze"
         .Add "MENU:TOOLS:OPTIONS", , "&Options..."
         
         ' Add Accounts:
         .Add "MENU:ACCOUNTS:LIST", , "&Account List"
         .Add "MENU:ACCOUNTS:BILLS", , "&Bills && Deposits"
         .Add "MENU:ACCOUNTS:MANAGER", , "&Online Service Manager"
         .Add "MENU:ACCOUNTS:SEP1", , , eSeparator
         .Add "MENU:ACCOUNTS:CASHFLOW", , "&Cash Flow"
         .Add "MENU:ACCOUNTS:CALENDAR", , "Ca&lendar"
         .Add "MENU:ACCOUNTS:SEP2", , , eSeparator
         .Add "MENU:ACCOUNTS:FAVOURITES", , "Fa&vourite Accounts"
         Set btn = .Add("MENU:ACCOUNTS:FAVOURITES:1", , "(No Favourites Yet)")
         btn.Enabled = False
         .Add "MENU:ACCOUNTS:SEP3", , , eSeparator
         .Add "MENU:ACCOUNTS:SETUP", , "Account &Setup"
         .Add "MENU:ACCOUNTS:CATEGORIES", , "Cate&gories"
         
         ' Add Investing
         .Add "MENU:INVESTING:PORTFOLIO", , "&Portfolio"
         .Add "MENU:INVESTING:ONLINE", , "&Online Investing Research"
         .Add "MENU:INVESTING:SEP1", , , eSeparator
         .Add "MENU:INVESTING:ANALYSIS", , "Portfolio &Analysis"
         .Add "MENU:INVESTING:ALLOCATION", , "A&sset Allocation"
         .Add "MENU:INVESTING:REPORTS", , "Investment &Reports"
         
         ' Add Planner
         .Add "MENU:PLANNER:LIFETIME", , "&Lifetime Planner"
         .Add "MENU:PLANNER:BUDGET", , "&Budget Planner"
         .Add "MENU:PLANNER:DEBT", , "&Debt Planner"
         .Add "MENU:PLANNER:INSURANCE", , "&Insurance Planner"
         .Add "MENU:PLANNER:SEP", , , eSeparator
         .Add "MENU:PLANNER:REPORTS", , "Planner &Reports"
         
         ' Add Taxes
         .Add "MENU:TAXES:ESTIMATOR", , "&Tax Estimator"
         .Add "MENU:TAXES:DEDUCTIONS", , "&Deduction Finder"
         Set btn = .Add("MENU:TAXES:WITHHOLDING", , "Tax &Withholding Estimator")
         btn.Enabled = False
         .Add "MENU:TAXES:LINE", , "Tax &Line Manager", , "Tax &Line Manager"
         .Add "MENU:TAXES:SEP1", , , eSeparator
         .Add "MENU:TAXES:SETTINGS", , "Tax &Settings"
         
         ' Add Shopping
         .Add "MENU:SHOPPING:CENTRE", , "&Shopping Centre"
         .Add "MENU:SHOPPING:BROKER", , "&Broker Centre"
         .Add "MENU:SHOPPING:BANKING", , "Ban&king Centre"
         
         ' Add Help
         .Add "MENU:HELP:CONTENTS", , "&Contents", , , vbKeyF1, 0
         .Add "MENU:HELP:SEP1", , , eSeparator
         .Add "MENU:HELP:WHATSTHIS", , "&What's This", , , vbKeyF1, vbShiftMask
         .Add "MENU:HELP:WEB", , "vbAccelerator on the &Web"
         .Add "MENU:HELP:REPAIR", , "&Repair"
         .Add "MENU:HELP:SEP2", , , eSeparator
         .Add "MENU:HELP:ABOUT", , "&About..."
         
         ' Add the toolbar buttons
         Set btn = .Add("TOOLBAR:BACK", ilsIcons.ItemIndex("BACK") - 1, "Back", , "Back", vbKeyLeft, vbAltMask)
         btn.Enabled = False
         btn.ShowDropDownInToolbar = True
         btn.ShowCaptionInToolbar = True
         
         Set btn = .Add("TOOLBAR:BACK:1", , "(None)")
         btn.Enabled = False
         
         Set btn = .Add("TOOLBAR:FORWARD", ilsIcons.ItemIndex("FORWARD") - 1, "Forward", , "Forward", vbKeyRight, vbAltMask)
         btn.Enabled = False
         btn.ShowDropDownInToolbar = True
         btn.ShowCaptionInToolbar = True
         
         Set btn = .Add("TOOLBAR:FORWARD:1", , "(None)")
         btn.Enabled = False
         
         Set btn = .Add("TOOLBAR:HOME", ilsIcons.ItemIndex("HOME") - 1, "Home", , "Go to your Home Page", vbKeyHome, vbAltMask)
         btn.ShowCaptionInToolbar = True
         
         .Add "TOOLBAR:SEP1", , , eSeparator
         Set btn = .Add("TOOLBAR:ADDRESS", , "Address", ePanel)
         btn.ShowCaptionInToolbar = True
         Set btn = .Add("TOOLBAR:ADDRESSCOMBO", , , ePanel)
         btn.PanelControl = cboAddress
         btn.PanelWidth = cboAddress.Width \ Screen.TwipsPerPixelX
         btn.ShowCaptionInToolbar = True
         Set btn = .Add("TOOLBAR:GO", ilsIcons.ItemIndex("GO"), "Go")
         btn.ShowCaptionInToolbar = True
         
         .Add "TOOLBAR:SEP2", , , eSeparator
         Set btn = .Add("TOOLBAR:ACCOUNTS", ilsIcons.ItemIndex("ACCOUNTS") - 1, "Account List")
         btn.ShowCaptionInToolbar = True

         Set btn = .Add("TOOLBAR:PORTFOLIO", ilsIcons.ItemIndex("PORTFOLIO") - 1, "Portfolio")
         btn.ShowCaptionInToolbar = True

         Set btn = .Add("TOOLBAR:BILLS", ilsIcons.ItemIndex("BILLS") - 1, "Bills && Deposits")
         btn.ShowCaptionInToolbar = True
         
         Set btn = .Add("TOOLBAR:WEB", ilsIcons.ItemIndex("VBACCELERATOR") - 1, "vbAccelerator")
         btn.ShowCaptionInToolbar = True

         Set btn = .Add("TOOLBAR:MORE", ilsIcons.ItemIndex("MORE") - 1, "More")
         btn.ShowCaptionInToolbar = True


         ' Add the status bar buttons:
         Set btn = .Add("STATUS:ONLINE", , "Online")
         btn.ShowCaptionInToolbar = True
         .Add "STATUS:SEP1", , , eSeparator
         Set btn = .Add("STATUS:UPDATES", 0, "Internet Updates")
         btn.ShowDropDownInToolbar = True
         btn.ShowCaptionInToolbar = True
         .Add "STATUS:SEP2", , , eSeparator
         Set btn = .Add("STATUS:STATUSTEXT")
         btn.Locked = True
         btn.ShowCaptionInToolbar = True
         
      End With
                  
   End With
   
End Sub

Private Sub setColour(ByVal lHue As Long)

   cmdBar(1).AdjustBackgroundImage lHue
   cmdBar(2).AdjustBackgroundImage lHue

   Dim r As Long, g As Long, b As Long
   HLSToRGB lHue, 85, 235, r, g, b
   Me.BackColor = RGB(r, g, b)
   
   Set picSideBar.Picture = cmdBar(0).AdjustImage(picRes.Picture, lHue, 85, 0.5)
   Dim sDate As String
   Dim lWidth As Long
   Dim lHeight As Long
   sDate = Format(Now, "mmmm dd, yyyy")
   lWidth = picSideBar.TextWidth(sDate)
   lHeight = picSideBar.TextHeight(sDate)
   picSideBar.CurrentX = 6
   picSideBar.CurrentY = (picSideBar.ScaleHeight - lHeight) \ 2
   picSideBar.Print sDate
   
End Sub

Private Sub addURL(ByVal sUrl As String)
   
   cboAddress.Tag = "ADDING"
   
   ' Detect if we already have this item:
   Dim i As Long
   Dim iIndex As Long
   iIndex = -1
   For i = 0 To cboAddress.ListCount - 1
      If (cboAddress.List(i) = sUrl) Then
         iIndex = i
         Exit For
      End If
   Next i
   
   If (iIndex = -1) Then
      cboAddress.AddItem sUrl, 0
      iIndex = 0
   End If
   cboAddress.ListIndex = iIndex
   
   If (iIndex = 0) Then
      cmdBar(0).Buttons("TOOLBAR:FORWARD").Enabled = False
   Else
      For i = 0 To iIndex - 1
         
      Next i
      cmdBar(0).Buttons("TOOLBAR:FORWARD").Enabled = True
   End If
   If (iIndex = cboAddress.ListCount - 1) Then
      cmdBar(0).Buttons("TOOLBAR:BACK").Enabled = False
   Else
      For i = iIndex + 1 To cboAddress.ListCount - 1
      Next i
      cmdBar(0).Buttons("TOOLBAR:BACK").Enabled = True
   End If
   
   cboAddress.Tag = ""
   
End Sub

Private Sub cboAddress_Click()
   If (cboAddress.Tag = "") Then
      web.Navigate2 cboAddress.Text
   End If
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
   If (KeyAscii = vbKeyReturn) Then
      Dim lIndex As Long
      Dim i As Long
      lIndex = -1
      For i = 0 To cboAddress.ListCount - 1
         If (cboAddress.List(i) = cboAddress.Text) Then
            lIndex = i
            Exit For
         End If
      Next i
      If (lIndex = -1) Then
         cboAddress.AddItem cboAddress.Text, 0
         lIndex = 0
      End If
      cboAddress.ListIndex = lIndex
   End If
End Sub

Private Sub cmdBar_ButtonClick(Index As Integer, btn As cButton)
   
   Debug.Print "Clicked", btn.Key
   Select Case btn.Key
   Case "TOOLBAR:ACCOUNTS"
      setColour 195
   Case "TOOLBAR:PORTFOLIO"
      setColour 45
   Case "TOOLBAR:BILLS"
      setColour 145
   Case "TOOLBAR:WEB", "HELP:MENU:WEB"
      setColour 75
      web.Navigate2 "http://vbaccelerator.com/"
   Case "TOOLBAR:HOME"
      web.Navigate2 App.Path & "\page.mht"
   Case "TOOLBAR:GO"
      cboAddress_KeyPress vbKeyReturn
   
   Case "MENU:HELP:ABOUT"
      Dim fA As New frmAbout
      fA.Acknowledgements = "This sample demonstrates the vbAccelerator CommandBars control using the MS Money rendering style."
      fA.Show vbModal, Me
   
   Case "MENU:FILE:EXIT"
      PostMessage Me.hWnd, WM_SYSCOMMAND, SC_CLOSE, 0

   End Select
   
End Sub

Private Sub cmdBar_RequestNewInstance(Index As Integer, ctl As Object)
   Dim lNewIndex As Long
   lNewIndex = cmdBar.UBound + 1
   Load cmdBar(lNewIndex)
   cmdBar(lNewIndex).Align = 0
   Set ctl = cmdBar(lNewIndex)
End Sub

Private Sub cmdBar_Resize(Index As Integer)
   If (Index = 0) Or (Index = 1) Then
      Form_Resize
   End If
End Sub

Private Sub cmdFont_Click()
Dim sFnt As New StdFont
   If (cmdBar(0).Font.Name = "Tahoma") Then
      sFnt.Name = "Times New Roman"
      sFnt.Size = 14
   Else
      sFnt.Name = "Tahoma"
      sFnt.Size = 8
   End If
   cmdBar(0).Font = sFnt
   cmdBar(1).Font = sFnt
   cmdBar(2).Font = sFnt
End Sub

Private Sub cmdStyle_Click()
    Dim i As Byte
    i = cmdBar(0).Style + 1
    If i > 3 Then i = 0
    cmdBar(0).Style = i
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler
      
   cmdBar(0).Redraw = False
   cmdBar(1).Redraw = False
   cmdBar(2).Redraw = False
   
   createButtons
   createCommandBars
   
   web.Navigate2 App.Path & "\page.mht"
   
   cmdBar(0).Style = eMoney
   'cmdBar(0).Style = eComCtl32
   'cmdBar(0).Style = eOfficeXP
   cmdBar(0).MainMenu = True
   cmdBar(0).ToolBar = cmdBar(0).CommandBars("MENU")
   
   cmdBar(1).BackgroundImage = picRes.Picture
   cmdBar(1).MenuImageList = ilsIcons16
   cmdBar(1).ToolbarImageList = ilsIcons
   cmdBar(1).ButtonTextPosition = eButtonTextBottom
   cmdBar(1).ToolBar = cmdBar(0).CommandBars("TOOLBAR")
   
   cmdBar(2).BackgroundImage = picRes.Picture
   cmdBar(2).MenuImageList = ilsIcons16
   cmdBar(2).ToolbarImageList = ilsIcons16
   cmdBar(2).ToolBar = cmdBar(0).CommandBars("STATUS")
   
   cmdBar(0).Redraw = True
   cmdBar(1).Redraw = True
   cmdBar(2).Redraw = True
   
   setColour 195
   
   Exit Sub
   
ErrorHandler:
   MsgBox "Error:" & Err.Description, vbExclamation
   Exit Sub
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Debug.Print cmdBar(0).MainMenu, cmdBar(1).MainMenu
End Sub

Private Sub Form_Resize()
Dim lTop As Long
Dim lLeft As Long
   On Error Resume Next
   lTop = cmdBar(1).top + cmdBar(1).Height
   lLeft = 128 * Screen.TwipsPerPixelX
   web.Move lLeft, lTop, Me.ScaleWidth - lLeft, Me.ScaleHeight - lTop - cmdBar(2).Height
   picSideBar.Move 0, lTop, lLeft, 30 * Screen.TwipsPerPixelY
   lTop = picSideBar.top + picSideBar.Height + 8 * Screen.TwipsPerPixelY
   lblInfo.Move 6 * Screen.TwipsPerPixelX, lTop, lLeft - 10 * Screen.TwipsPerPixelX, Me.ScaleHeight - lTop
End Sub

Private Sub web_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
   '
   addURL URL
   '
End Sub

Private Sub web_NewWindow2(ppDisp As Object, Cancel As Boolean)
   '
   Cancel = True
   '
End Sub

Private Sub web_StatusTextChange(ByVal Text As String)
   cmdBar(2).Buttons("STATUS:STATUSTEXT").Caption = Text
End Sub

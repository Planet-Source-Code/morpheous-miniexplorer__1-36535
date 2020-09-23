VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Mini Explorer"
   ClientHeight    =   6030
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8460
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   582
      ButtonWidth     =   609
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Up"
            Object.ToolTipText     =   "Go Up One Directory"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Object.ToolTipText     =   "Show File or Directory Properties"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Back"
            Object.ToolTipText     =   "History"
            ImageIndex      =   7
            Style           =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   6360
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":57E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":C044
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":E7F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":10FA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":17242
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1D4DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar BtnGo 
      Height          =   330
      Left            =   8160
      TabIndex        =   2
      Top             =   360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   582
      ButtonWidth     =   609
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7080
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":23776
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":23D10
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":23E22
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":23F34
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2507E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":261C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":26ED2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox CboPath 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   8175
   End
   Begin MSComctlLib.ListView List 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList2"
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Attributes"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu MnuFileNewWindow 
         Caption         =   "Open in &New Window"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuFileRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   ^R
      End
      Begin VB.Menu MnuFileRun 
         Caption         =   "Run/&Execute"
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuFileProperties 
         Caption         =   "Proper&ties"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuFileSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFileClose 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ######################################################################
'#                  Author: Eric Cogen                                  #
'#                  Date: July/03/2002                                  #
'#                  Website: http://www27.brinkster.com/countoia/       #
'#                  Email: ericcogen@yahoo.com                          #
'#                                                                      #
'# If you use any or all of this code please include me in your header. #
'# Thanks                                                               #
' ######################################################################

':....References: Make sure to include these in your project....:
'  :....Visual Basic For Applications....:
'  :....Visual Basic runtime objects and procedures....:
'  :....OLE Automation....:
'  :....Microsoft Scripting Runtime....:
'Project..References From the Main Menu Or Alt+P+N

Option Compare Text
'API Help Courtesy of http://www.allapi.net/agnet/appdown.htm ... Great App Get it!
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef S As SHELLEXECUTEINFO) As Long

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type


Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Const MAX_PATH = 260

Const CSIDL_SENDTO = &H9
Const CSIDL_FONTS = &H14
Const CSIDL_TEMPLATES = &H15
Const CSIDL_STARTMENU = &HB
Const CSIDL_DESKTOPDIRECTORY = &H10
Const CSIDL_FAVORITES = &H6
Const CSIDL_PRINTERS = &H4
Const CSIDL_PROGRAMS = &H2
Const CSIDL_DESKTOP = &H0

Const CSIDL_NETWORK = &H12
Const CSIDL_NETHOOD = &H13
Const CSIDL_PERSONAL = &H5

Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOW = 5
Private Const SEE_MASK_INVOKEIDLIST = &HC

Dim FSO As New FileSystemObject
Dim D As Drive
Dim Fol As Folder
Dim Fil As File

Dim LI As ListItem
Dim TempNode As Node

Dim StopActions As Boolean

Private Sub BtnGo_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error_Trap

   Call ListItems(CboPath.Text)
   
Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Sub BtnGo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Error_Trap

 BtnGo.Buttons(1).ToolTipText = "Go to " & AppendBS(CboPath.Text)

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Sub CboPath_Change()
On Error GoTo Error_Trap

If CboPath.Text = "Network N" Then
   CboPath.Text = "Network Neighborhood"
ElseIf CboPath.Text = "Control P" Then
   CboPath.Text = "Control Panel"
ElseIf CboPath.Text = "My Doc" Then
   CboPath.Text = "My Documents"
ElseIf CboPath.Text = "My Com" Then
   CboPath.Text = "My Computer"
End If
CboPath.SelStart = Len(CboPath.Text)

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Sub CboPath_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Error_Trap

If KeyCode = vbKeyReturn Then
   Call ListItems(CboPath.Text)
End If

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo Error_Trap
Dim Com As String

Com = Command ':....incoming commandline switch....:
 
':....maximize switch takes precedence over center....:
 
If InStr(1, Com, "/m") > 0 Then ':....switch to maximize window on load....:
   Me.WindowState = vbMaximized
   Com = Replace$(Com, "/c", "", , , vbTextCompare)
   Com = Replace$(Com, "/m", "", , , vbTextCompare)
End If

If InStr(1, Com, "/c") > 0 Then ':....switch to center window on load....:
   Me.Left = (Screen.Width \ 2) - (Me.Width \ 2)
   Me.Top = (Screen.Height \ 2) - (Me.Height \ 2)
   Com = Replace$(Com, "/c", "", , , vbTextCompare)
   Com = Replace$(Com, "/m", "", , , vbTextCompare)
End If

Me.Show 'Force window to show
Form_Resize

List.ColumnHeaders(1).Width = List.Width \ 3
List.ColumnHeaders(2).Width = List.Width \ 3
List.ColumnHeaders(3).Width = List.Width \ 3

Com = Replace$(Com, Chr$(39), "", , , vbTextCompare) '= Quote ..: Replace it with nothing
Com = Replace$(Com, Chr$(34), "", , , vbTextCompare) '=Single Quote ..: Replace it with nothing

If Trim$(Com) <> "" Then
   Call ListItems(Com)
Else
   Call ListItems(GWD)
End If

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo Error_Trap

If Me.WindowState <> vbMinimized Then
   CboPath.Width = Me.Width - 100 - BtnGo.Width
   List.Width = Me.Width - 100
   BtnGo.Left = Me.Width - BtnGo.Width - 100
   List.Height = Me.Height - Toolbar1.Height - CboPath.Height - 900
   CboPath.Left = 0
End If

Exit Sub
Error_Trap: 'Let it error out and continue '''Not very good but it works
 Err.Clear
Exit Sub
End Sub

Sub ListItems(Path As String)
On Error GoTo Error_Trap

Dim I As Integer

Screen.MousePointer = vbHourglass

Call ATC(Path)

List.ListItems.Clear

If Trim$(Path) = "My Computer" Then
   Screen.MousePointer = vbDefault
   Call GetMyComputer
   Exit Sub
ElseIf Trim$(Path) = "My Documents" Then
   Screen.MousePointer = vbDefault
   Call GetMyDocuments
   Exit Sub
ElseIf Trim$(Path) = "Control Panel" Then
   Screen.MousePointer = vbDefault
   Call GetControlPanel
   Exit Sub
ElseIf Trim$(Path) = "Network Neighborhood" Then
   Screen.MousePointer = vbDefault
   Call GetNetwork
   Exit Sub
End If

Set Fol = FSO.GetFolder(Path)

For Each Fol In Fol.SubFolders
    If Not StopActions Then
       DoEvents
       Set LI = List.ListItems.Add(, , Fol.Name, 1, 1)
       LI.ListSubItems.Add , , Fol.Type
       LI.ListSubItems.Add , , Fol.Attributes
    End If
Next

Set Fol = FSO.GetFolder(Path)

For Each Fil In Fol.Files
    If Not StopActions Then
       DoEvents
       Set LI = List.ListItems.Add(, , Fil.Name, 2, 2)
       LI.ListSubItems.Add , , Fil.Type
       LI.ListSubItems.Add , , Fil.Attributes
    End If
Next

Set Fil = Nothing
Set Fol = Nothing
Screen.MousePointer = vbDefault
List.Sorted = False
StopActions = False

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Screen.MousePointer = vbDefault
 Err.Clear
Exit Sub
End Sub

Sub GoUp(Path As String)
On Error GoTo Error_Trap
Dim WIS As Long

WIS = InStrRev(Path, "\", Len(Path) - 1, vbTextCompare)

If WIS > 0 Then
   Call ListItems(Left(Path, WIS))
Else
   Call GetMyComputer
End If

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Function ATC(Path As String) As String 'Add To ComboList
On Error GoTo Error_Trap

Dim I As Integer
Dim ATP As Boolean

ATP = True

For I = 0 To CboPath.ListCount - 1
    If CboPath.List(I) = Path Then
       ATP = False
    End If
Next I

If ATP Then
   CboPath.AddItem UCase(Left(Path, 1)) & Right(Path, Len(Path) - 1)
   Toolbar1.Buttons(3).ButtonMenus.Add , , UCase(Left(Path, 1)) & Right(Path, Len(Path) - 1)
End If

CboPath.Text = UCase(Left(Path, 1)) & Right(Path, Len(Path) - 1)
Me.Caption = "MiniExplorer - " & UCase(Left(Path, 1)) & Right(Path, Len(Path) - 1) 'Capitalize the first letter

Exit Function
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Function
End Function

Function GWD() As String 'Get Windows Directory
On Error GoTo Error_Trap
Dim Path As String, StrSave As String
 
 StrSave = String(200, Chr$(0))
 Path = Left$(StrSave, GetWindowsDirectory(StrSave, Len(StrSave)))
 GWD = Path

Exit Function
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Function
End Function

Private Sub List_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo Error_Trap 'Email me if you know of a better one please

If List.SortKey = ColumnHeader.Index - 1 Then
   If List.SortOrder = lvwAscending Then
      List.SortOrder = lvwDescending
   Else
      List.SortOrder = lvwAscending
   End If
Else
   List.SortOrder = lvwAscending
   List.SortKey = ColumnHeader.Index - 1
End If

List.Sorted = True

Exit Sub
Error_Trap:
 Err.Clear
Exit Sub
End Sub

Private Sub List_DblClick()
On Error GoTo Error_Trap

If List.SelectedItem.SmallIcon = 1 Then 'Folder
   Call ListItems(AppendBS(CboPath.Text) & List.SelectedItem.Text)
Else
  If List.SelectedItem.SmallIcon <> 2 Then 'Not a file so either one of the drives icons
   Call ListItems(AppendBS(List.SelectedItem.Text))
  End If
End If

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Sub List_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Error_Trap

If KeyCode = 27 Then 'Escape Key
   StopActions = True
   Exit Sub
End If

StopActions = False

If KeyCode = 116 Then 'F5 Key...: Typically the refresh key
   Call ListItems(CboPath.Text)
End If

If KeyCode = vbKeyReturn Then 'Return Or Enter Key
   If List.SelectedItem.SmallIcon = 1 Then
      Call ListItems(AppendBS(CboPath.Text) & List.SelectedItem.Text)
   End If
Else
   If vbAltMask Then 'Alt + LeftArrow
      If KeyCode = 37 Then
         Call GoUp(CboPath.Text)
      ElseIf KeyCode = 39 Then 'Alt + RightArrow
             If List.SelectedItem.SmallIcon = 1 Then
                Call ListItems(AppendBS(CboPath.Text) & List.SelectedItem.Text)
             End If
      End If
   End If
End If

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Sub List_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Error_Trap

If Button = vbRightButton Then
   PopupMenu MnuFile
End If

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Sub List_OLECompleteDrag(Effect As Long)
On Error GoTo Error_Trap

 Call ListItems(CboPath.Text)

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Sub List_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Error_Trap 'Allow drag and drop into and from windows

Dim DroppedItem As Variant

Screen.MousePointer = vbHourglass

For Each DroppedItem In Data.Files
    DoEvents 'During long loops its always a good idea to return control to windows
    If FSO.FileExists(DroppedItem) Then
       If Shift <> 2 Then
          FSO.MoveFile DroppedItem, AppendBS(CboPath.Text)
       Else
          FSO.CopyFile DroppedItem, AppendBS(CboPath.Text)
       End If
    ElseIf FSO.FolderExists(DroppedItem) Then
           If Shift <> 2 Then
              FSO.MoveFolder DroppedItem, AppendBS(CboPath.Text)
           Else
              FSO.CopyFolder DroppedItem, AppendBS(CboPath.Text)
           End If
    End If
Next

Screen.MousePointer = vbDefault

Call ListItems(CboPath.Text)

Exit Sub
Error_Trap:
 Screen.MousePointer = vbDefault
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Sub List_OLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)
On Error GoTo Error_Trap

For Each LI In List.ListItems
    DoEvents
    If LI.Selected = True Then
       Data.Files.Add AppendBS(CboPath.Text) & LI.Text
    End If
Next

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Sub List_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
On Error GoTo Error_Trap

AllowedEffects = vbDropEffectMove Or vbDropEffectCopy
Data.SetData , 15

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Sub MnuFileClose_Click()
On Error GoTo Error_Trap

 Unload Me 'Never Ever Ever use the End keyword!

Exit Sub
Error_Trap:
 End
 Err.Clear
Exit Sub
End Sub

Private Sub MnuFileNewWindow_Click()
On Error GoTo Error_Trap

Dim I As Integer

For I = 1 To List.ListItems.Count
    DoEvents
    If List.ListItems.Item(I).Selected Then
       If List.ListItems.Item(I).SmallIcon = 1 Then
          If Right$(Trim$(App.Path), 1) = "\" Then
             Shell App.Path & App.EXEName & ".exe " & AppendBS(CboPath.Text) & List.ListItems.Item(I).Text, vbNormalFocus
          Else
             Shell App.Path & "\" & App.EXEName & ".exe " & AppendBS(CboPath.Text) & List.ListItems.Item(I).Text, vbNormalFocus
          End If
       End If
     End If
Next I

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Sub MnuFileOpen_Click()
On Error GoTo Error_Trap

If List.SelectedItem.SmallIcon = 1 Then
      Call ListItems(AppendBS(CboPath.Text) & List.SelectedItem.Text)
End If

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Sub MnuFileProperties_Click()
On Error GoTo Error_Trap

 Call ShowProps(AppendBS(CboPath.Text) & List.SelectedItem.Text)

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Sub MnuFileRefresh_Click()
On Error GoTo Error_Trap

 StopActions = False
 Call ListItems(CboPath.Text)

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Sub MnuFileRun_Click()
On Error GoTo Error_Trap
Dim Ret As Long

If List.SelectedItem.SmallIcon <> 1 Then
   Ret = ShellExecute(Me.hWnd, vbNullString, AppendBS(CboPath.Text) & List.SelectedItem.Text, vbNullString, "C:\", SW_SHOWNORMAL)
   If Ret = 31 Then 'Show the Open With Dialog Box
      ShellExecute GetDesktopWindow, vbNullString, "RUNDLL32.EXE", "shell32.dll,OpenAs_RunDLL " & List.SelectedItem.Text, AppendBS(CboPath.Text), vbNormalFocus
   End If
Else
   Call ListItems(AppendBS(CboPath.Text) & List.SelectedItem.Text)
End If

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error_Trap

If Button.Key = "Up" Then
   Call GoUp(CboPath.Text)
ElseIf Button.Key = "Properties" Then
       Call ShowProps(AppendBS(CboPath.Text) & List.SelectedItem.Text)
End If

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo Error_Trap

 Call ListItems(ButtonMenu.Text)

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Sub ShowProps(Path As String)
On Error GoTo Error_Trap

Dim ShInfo As SHELLEXECUTEINFO

With ShInfo
 .cbSize = LenB(ShInfo)
 .lpFile = Path
 .nShow = SW_SHOW
 .fMask = SEE_MASK_INVOKEIDLIST
 .lpVerb = "Properties"
End With

ShellExecuteEx ShInfo

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Function AppendBS(Path As String) As String 'Append Backslash (\) to end of path
On Error GoTo Error_Trap

If Right$(Trim$(Path), 1) = "\" Then
   AppendBS = Path
Else
   AppendBS = Path & "\"
End If

Exit Function
Error_Trap:
 Err.Clear
Exit Function
End Function

Private Sub GetMyComputer()
On Error GoTo Error_Trap

Screen.MousePointer = vbHourglass

List.ListItems.Clear

Call ATC("My Computer")

For Each D In FSO.Drives
    DoEvents 'Might need this in case of latency issues with networked drives!
    Set LI = List.ListItems.Add(, , D.DriveLetter & ":\", , D.DriveType + 2)
        If D.IsReady Then
           LI.ListSubItems.Add , , D.FileSystem
        Else
           LI.ListSubItems.Add , , "Insert Media..."
        End If
Next
Screen.MousePointer = vbDefault

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Sub GetMyDocuments()
On Error GoTo Error_Trap

 Call ListItems(GetSpecialfolder(CSIDL_PERSONAL)) 'Returns something like "C:\Documents and Settings\[% Current User %]\My Documents"

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Sub GetControlPanel()
On Error GoTo Error_Trap

Dim sSave As String, Ret As Long

List.ListItems.Clear

Screen.MousePointer = vbHourglass
sSave = Space(255)
Ret = GetSystemDirectory(sSave, 255)
sSave = Left$(sSave, Ret)

Set Fol = FSO.GetFolder(sSave)

For Each Fil In Fol.Files
    DoEvents
    If Right$(Fil.Name, 3) = "cpl" Then 'Only show the Control Panel Applets
       Set LI = List.ListItems.Add(, , Fil.Name, , 2)
           LI.ListSubItems.Add , , Fil.Type
           LI.ListSubItems.Add , , Fil.Attributes
    End If
Next
Screen.MousePointer = vbDefault

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Sub GetNetwork()
On Error GoTo Error_Trap

Call ListItems(GetSpecialfolder(CSIDL_NETHOOD)) 'Untested ...: No Network
'Or
'Call ListItems(GetSpecialfolder(CSIDL_NETWORK))'Returns blank 'Try on your systems and let me know.

Exit Sub
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Sub
End Sub

Private Function GetSpecialfolder(CSIDL As Long) As String
On Error GoTo Error_Trap

Dim R As Long
Dim IDL As ITEMIDLIST
    
    R = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    
    If R = NOERROR Then
        Path$ = Space$(512)
        R = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
        GetSpecialfolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""

Exit Function
Error_Trap:
 MsgBox Err.Description, vbInformation
 Err.Clear
Exit Function
End Function

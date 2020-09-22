VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   7875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change Path"
      Height          =   315
      Left            =   6600
      TabIndex        =   0
      Top             =   75
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7125
      Top             =   5625
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   "desc"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":059C
            Key             =   "blank"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0936
            Key             =   "asc"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   465
      Left            =   6600
      TabIndex        =   5
      Top             =   450
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save List"
      Enabled         =   0   'False
      Height          =   465
      Left            =   5325
      TabIndex        =   4
      Top             =   450
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete Checked Items"
      Height          =   465
      Left            =   3225
      TabIndex        =   3
      Top             =   450
      Width           =   2040
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "&Select All"
      Height          =   240
      Left            =   100
      TabIndex        =   1
      Top             =   675
      Width           =   1215
   End
   Begin VB.CommandButton cmdGetList 
      Caption         =   "&Get List"
      Height          =   465
      Left            =   1950
      TabIndex        =   2
      Top             =   450
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwList 
      Height          =   5040
      Left            =   75
      TabIndex        =   6
      Top             =   1275
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   8890
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "List #"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Project Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Path"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Check the item(s) to delete, Delete them, then Save List to finalize the process."
      Height          =   195
      Left            =   1155
      TabIndex        =   8
      Top             =   1050
      Width           =   5565
   End
   Begin VB.Label lblRegPath 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Registry Path: "
      Height          =   195
      Left            =   5460
      TabIndex        =   7
      Top             =   150
      Width           =   1035
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  'registry key path must have a final backslash
  Const DefaultRegKey = "HKCU\Software\Microsoft\Visual Basic\6.0\RecentFiles\"
  Dim FSO As FileSystemObject
  
Private Sub chkAll_Click()
  'if chkAll is checked, then check all items in the list view, otherwise
  'uncheck all items in the listview
  If chkAll Then
    For J = 1 To lvwList.ListItems.Count
      lvwList.ListItems(J).Checked = True
    Next J
  ElseIf Not chkAll Then
    For J = 1 To lvwList.ListItems.Count
      lvwList.ListItems(J).Checked = False
    Next J
  End If
End Sub

Private Sub cmdChange_Click()
  'allow the user to enter a different registry path.
  Dim tmpPath As String
  tmpPath = InputBox("Enter a new Registry Path", "Registry Path", "HKCU\Software\Microsoft\Visual Basic\6.0\RecentFiles\")
  If tmpPath <> "" Then
    lblRegPath.Caption = tmpPath
  End If
End Sub

Private Sub cmdDelete_Click()
  Dim tmpSortKey As Integer     'hold the current sort key
  Dim tmpSortOrder As Integer   'hold the current sort order
  'we have to sort the list back to the original order before we remove an item from the list
  'SEE COMMENTS IN cmdSave_Click() FOR MORE INFORMATIN ABOUT THIS
  With lvwList
    tmpSortKey = .SortKey
    tmpSortOrder = .SortOrder
    .SortKey = 0
    .SortOrder = lvwAscending
  End With
  'delete the items from the list and renumber the remaining items
  For J = lvwList.ListItems.Count To 1 Step -1
    If lvwList.ListItems(J).Checked Then
      lvwList.ListItems.Remove J
      'enable save button since we have actually deleted an item
      cmdSave.Enabled = True
    End If
  Next J
  For J = 1 To lvwList.ListItems.Count
    lvwList.ListItems(J).Text = Format(J, "00")
  Next J
  'once we have removed the item(s) and renumbered everything, we can change the sorting
  'back to what the user had.
  With lvwList
    .SortKey = tmpSortKey
    .SortOrder = tmpSortOrder
  End With
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdGetList_Click()
  On Error GoTo ErrHandler
  Set FSO = New FileSystemObject
  
  'temp variable to hold the value of the registry item
  Dim tmpValue As String
  'clear the list
  lvwList.ListItems.Clear
  'clear the chkAll box
  chkAll = False
  'create the shell scripting object
  CreateObj
  'cycle through the registry values, we know there are at a maximum 1 to 50 items
  For J = 1 To 50
    'get the registry value
    tmpValue = WSRead(lblRegPath.Caption & J)
    If tmpValue <> "" Then
      'if tmpValue is not blank, we have gotten a registry key value,
      'now we add the index number to the list view as a list item.
      'Then, we add the first sub item to this list item which is the name of the project name
      'now we add the second sub item to this list item which is the path to the project file
      lvwList.ListItems.Add , , Format(J, "00")
      'pull out the project name by starting one character AFTER the last backslash
      'and get the length of characters to pull out by counting between the last backslash
      'and the period that notes the file extention.
      lvwList.ListItems(J).ListSubItems.Add , , Mid(tmpValue, InStrRev(tmpValue, "\") + 1, (InStrRev(tmpValue, ".")) - (InStrRev(tmpValue, "\") + 1))
      lvwList.ListItems(J).ListSubItems.Add , , tmpValue
      'check to see if the project file still exists in the noted path
      If Not FSO.FileExists(tmpValue) Then
        With lvwList
          .ListItems(J).ForeColor = vbRed
          .ListItems(J).ListSubItems(1).ForeColor = vbRed
          .ListItems(J).ListSubItems(2).ForeColor = vbRed
        End With
      End If
    End If
  Next J
  
  'resize the columns in the list view to show all content
ResizeColumns:
  For J = 1 To lvwList.ColumnHeaders.Count
    SizeColumn J, lvwList
  Next J
  Set FSO = Nothing
  Exit Sub
ErrHandler:
  'do nothing: two possible causes of an error here -
  ' 1) the key path is incorrect - fix - check registry for the right path
  ' 2) there are not 50 keys in the registry - fix - exit this routine as
  '    we have found all of the registry values there are
  'BUT FIRST RESIZE THE COLUMNS
  GoTo ResizeColumns
End Sub

Private Sub cmdSave_Click()
  'first we need to set the sorting back to the first column (numbers)
  'and make sure the sort order is ascending, this will save the list
  'just as it was when VB made it.
  '**************************************
  'IF YOU WANT TO SAVE THE LIST IN ANOTHER SORTED ORDER COMMENT OUT THE
  'WITH STATEMENT. THE LIST WILL BE SAVED IN THE CURRENT ORDER IT
  'IS SORTED IN.
  '**************************************
  With lvwList
    .SortKey = 0
    .SortOrder = lvwAscending
  End With
  'first we have to delete all current registry values
  'by deleting the recent files folder itself
  WSDelete lblRegPath
  'now we write all of the remaining values back into the registry
  'for future use by VB.
  For J = 1 To lvwList.ListItems.Count
    WSWrite lblRegPath.Caption & J, lvwList.ListItems(J).ListSubItems(2).Text
  Next J
  lvwList_ColumnClick lvwList.ColumnHeaders(1)

  'disable button until further changes are made
  cmdSave.Enabled = False
End Sub

Private Sub Form_Load()
  'load the default registry key that contains the VB Recent List
  lblRegPath.Caption = DefaultRegKey
  'initialize the list
  cmdGetList_Click
  'set some list view properties
  With lvwList
    .ColumnHeaderIcons = ImageList1 'icons to use for sorting
    .ColumnHeaders(1).Icon = "asc"  'set first column's icon
    .Sorted = True
    .SortKey = 0                    'sorted on first column
  End With
End Sub

Private Sub Form_Resize()
  'move the  the list view accordingly
  With lvwList
    .Left = Me.ScaleLeft
    .Width = Me.ScaleWidth
    .Height = Me.ScaleHeight - .Top
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  DeleteObj
  Set frmMain = Nothing
End Sub

Private Sub lvwList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Static srtdColumn As Integer                          'tracks what column is sorted
  lvwList.ColumnHeaders(srtdColumn + 1).Icon = "blank"  'clear the icon
  With lvwList
  .SortKey = ColumnHeader.Index - 1                     'change sorted column
  If srtdColumn = ColumnHeader.Index - 1 Then
    If .SortOrder = lvwAscending Then
      .SortOrder = lvwDescending                        'change sort order
      ColumnHeader.Icon = "desc"                        'and icon
    Else
      .SortOrder = lvwAscending
      ColumnHeader.Icon = "asc"
    End If
    srtdColumn = ColumnHeader.Index - 1                 'update what column is sorted
  Else
    .SortOrder = lvwAscending
    ColumnHeader.Icon = "asc"
    srtdColumn = ColumnHeader.Index - 1
  End If
  End With
End Sub

Private Sub lvwList_ItemClick(ByVal Item As MSComctlLib.ListItem)
'UN COMMENT THIS BLOCK IF YOU WANT TO BE ABLE TO CLICK ANYWHERE ON THE ITEM AND CHECK IT
'  'check/uncheck the item clicked
'  If Not Item.Checked Then
'    Item.Checked = True
'  Else
'    Item.Checked = False
'  End If
End Sub

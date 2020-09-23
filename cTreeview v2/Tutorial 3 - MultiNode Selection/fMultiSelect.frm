VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fMultiSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tutorial 3 : Multi-Node Selection"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fMultiSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOption 
      Caption         =   "O&ptions: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   4410
      TabIndex        =   9
      Top             =   5145
      Width           =   4845
      Begin VB.CheckBox chkOption 
         Appearance      =   0  'Flat
         Caption         =   "No Default Selection"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   420
         TabIndex        =   13
         Top             =   1215
         Value           =   1  'Checked
         Width           =   2805
      End
      Begin VB.CheckBox chkOption 
         Appearance      =   0  'Flat
         Caption         =   "No Clear On Space Click"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   420
         TabIndex        =   12
         Top             =   915
         Width           =   2805
      End
      Begin VB.CheckBox chkOption 
         Appearance      =   0  'Flat
         Caption         =   "Bold Selected Nodes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   420
         TabIndex        =   11
         Top             =   615
         Width           =   2805
      End
      Begin VB.CheckBox chkOption 
         Appearance      =   0  'Flat
         Caption         =   "Use Default selection Colours"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   420
         TabIndex        =   10
         Top             =   315
         Value           =   1  'Checked
         Width           =   2805
      End
   End
   Begin VB.ListBox lstDialog 
      BackColor       =   &H80000018&
      Height          =   4740
      Left            =   6000
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   315
      Width           =   3255
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "T&ransfer ->>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   4515
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "&Toggle Node"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   4515
      TabIndex        =   3
      Top             =   1260
      Width           =   1215
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "&Clear Node"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   4515
      TabIndex        =   1
      Top             =   210
      Width           =   1215
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "&Select Node"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   4515
      TabIndex        =   5
      Top             =   2310
      Width           =   1215
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "Select &All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   4515
      TabIndex        =   6
      Top             =   2730
      Width           =   1215
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "Clear A&ll"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   4530
      TabIndex        =   2
      Top             =   630
      Width           =   1215
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "To&ggle All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   4515
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CheckBox chkAutoList 
      Appearance      =   0  'Flat
      Caption         =   "Auto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4935
      TabIndex        =   8
      Top             =   3780
      Value           =   1  'Checked
      Width           =   645
   End
   Begin MSComctlLib.ImageList ilDialog 
      Left            =   4830
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483628
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMultiSelect.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMultiSelect.frx":8A3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMultiSelect.frx":ECD6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwDialog 
      Height          =   6630
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   11695
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   1
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblDialog 
      Caption         =   "Selected Nodes:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6885
      TabIndex        =   14
      Top             =   0
      Width           =   2325
   End
End
Attribute VB_Name = "fMultiSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'
' Form Name:    fMultiSelect [Tutorial 3]
' Author:       Graeme Grant        (a.k.a. Slider)
' Date:         23/03/2002
' Version:      01.00.00
' Description:  Test/Demo TreeView Handler
' Edit History: 01.00.00 23/03/2002 Initial Release
'
'===========================================================================

Option Explicit

Private Const cSELFORECOLOR As Long = vbYellow
Private Const cSELBACKCOLOR As Long = vbRed

#If NODLL = 0 Then
    Private WithEvents moTree As vbTree.cTreeView
Attribute moTree.VB_VarHelpID = -1
    Private miMultiSelect     As vbTree.iMultiSelect
#Else
    Private WithEvents moTree As cTreeView
Attribute moTree.VB_VarHelpID = -1
    Private miMultiSelect     As iMultiSelect
#End If

Private Enum eCommand
    eClear = 0
    eClearAll = 1
    eToggle = 2
    eToggleAll = 3
    eSelect = 4
    eSelectAll = 5
    eTransfer = 6
End Enum

Private Enum eCheck
    eDefColor = 0
    eSelBold = 1
    eNoCLear = 2
    eDefSel = 3
End Enum

'===========================================================================
' Form Events
'
Private Sub Form_KeyPress(KeyAscii As Integer)
    '
    '## Make Return/Enter key act like the Tab key...
    '
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys "{TAB}"
    End Select

End Sub

Private Sub Form_Load()

#If NODLL = 0 Then
    Set moTree = New vbTree.cTreeView   '## Used to manage the TreeView
#Else
    Set moTree = New cTreeView          '## Used to manage the TreeView
#End If

    Set miMultiSelect = moTree

    With moTree
        '
        '## Hook treeview control
        '
        .HookCtrl tvwDialog, [Multi Select]
        '
        '## Set TreeView features
        '
        With .Ctrl
            .Style = tvwTreelinesPlusMinusPictureText
            .LineStyle = tvwRootLines
            .Indentation = 10
            .ImageList = ilDialog
            .FullRowSelect = False
            .HideSelection = False
            .HotTracking = True
            '
            '## Build TreeView data
            '
            .Visible = False
            pInitData
            .Visible = True
            '
            '## Show focus rectangle over first node but don't select
            '
            With .Nodes(1)
                .Selected = True
                .Selected = False
            End With
        End With
    End With

End Sub

'===========================================================================
' Form Control Events
'
Private Sub chkOption_Click(Index As Integer)

    With miMultiSelect
        Select Case Index
            Case eDefColor
                Select Case chkOption(Index).Value
                    Case vbUnchecked
                        .SelBackColor = cSELBACKCOLOR
                        .SelForeColor = cSELFORECOLOR
                    Case vbChecked
                        .SelBackColor = vbHighlight
                        .SelForeColor = vbHighlightText
                End Select

            Case eSelBold
                .SelBold = CBool(chkOption(Index).Value)
            Case eNoCLear
                .NoClearOnSpaceClick = CBool(chkOption(Index).Value)

            Case eDefSel
                .NoDefaultSel = CBool(chkOption(Index).Value)

        End Select
    End With
    tvwDialog.SetFocus

End Sub

Private Sub cmdDialog_Click(Index As Integer)

    Dim oNode As MSComctlLib.Node

    Select Case Index
        Case eClear
            With miMultiSelect
                .ClearSelection .FocusNode
                pCmdsEnable Not (.FocusNode Is Nothing)
            End With

        Case eClearAll
            With miMultiSelect
                .ClearSelection
                pCmdsEnable Not (.FocusNode Is Nothing)
            End With

        Case eToggle
            With miMultiSelect
                .ToggleSelection .FocusNode
                pCmdsEnable Not (.FocusNode Is Nothing)
            End With

        Case eToggleAll
            With miMultiSelect
                .ToggleSelection
                pCmdsEnable Not (.FocusNode Is Nothing)
            End With

        Case eSelect
            With miMultiSelect
                .SelectAllNodes .FocusNode
                pCmdsEnable Not (.FocusNode Is Nothing)
            End With

        Case eSelectAll
            With miMultiSelect
                .SelectAllNodes
                pCmdsEnable Not (.FocusNode Is Nothing)
            End With

        Case eTransfer
            With lstDialog
                .Visible = False
                .Clear
                For Each oNode In miMultiSelect
                    .AddItem oNode.Text
                Next
                .Visible = True
            End With

    End Select
    tvwDialog.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moTree = Nothing
End Sub

'===========================================================================
' Form Control Events
'
Private Sub moTree_NodeClick(ByVal Node As MSComctlLib.Node)
    With miMultiSelect
        pCmdsEnable Not (.FocusNode Is Nothing)
    End With
End Sub

Private Sub moTree_SelChange()
    If chkAutoList.Value Then
        cmdDialog_Click eTransfer
    End If
End Sub

'===========================================================================
' Internal Functions
'
Private Sub pInitData()

    With moTree
        .NodeAdd , , "A", "Basic Functions", 2, 1, , , , True, , True
        .NodeAdd , , "B", "Drag and Drop", 2, 1
        .NodeAdd , , "C", "MultiSelection", 2, 1
        .NodeAdd , , "D", "Load On Demand", 2, 1
        .NodeAdd , , "E", "ADO Integration", 3, 3

        Dim lLoop As Long
        .NodeAdd , , "X1", "Node Item 1", 2, 1
        For lLoop = 2 To 50
            .NodeAdd tvwDialog.Nodes("X" + CStr(lLoop - 1)), tvwChild, "X" + CStr(lLoop), "Node Item " + CStr(lLoop), 2, 1
        Next
    End With

End Sub

Private Sub pCmdsEnable(Mode As Boolean)
    cmdDialog(eClear).Enabled = Mode
    cmdDialog(eToggle).Enabled = Mode
    cmdDialog(eSelect).Enabled = Mode
End Sub

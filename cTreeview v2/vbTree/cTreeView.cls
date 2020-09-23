VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===========================================================================
'
' Class Name:  cTreeView
' Author:      Graeme Grant        (a.k.a. Slider)
' Date:        23/03/2001
' Version:     02.01.00
' Description: Advanced TreeView Handler
' History:     Too big to put here - See cTreeviewHistory.Txt file
' Compiling:   DLL : set class Instancing = 5 - MultiUse
'              APP : set class Instancing = 1 - Private
' Notes:       NODLL conditional compilation flag is used in the project's
'              'Conditional Compiling Arguments' property (see
'              tutorial?_NODLL.PRJ properties to see usage).
'
' Original Authors: Credit where it's due:- MultiNode Selection -
'                   www.aboutvb.de; TrueDragImage, SetOverlayImage,
'                   CutIconState, Locked, & Enabled - Brad Martinez;
'                   FirstVisibleNode - www.vb2themax.com; ScrollView -
'                   Microsoft (Articles Q177743 & Q244312).
'
'===========================================================================

Option Explicit

'===========================================================================
' Debugging... Saves adding the debug statements to the form events
'
#Const DEBUGMODE = 1                    '## 0=No debug
                                        '   1=debug
#Const MOUSEEVENTS = 1                  '## 0=No mouse events
                                        '   1=Mouse Up & Mouse Down
                                        '   2=All Mouse events
#If DEBUGMODE = 1 Then
    Private dbgCtrlName  As String
#End If

'===========================================================================
' Public: Variables and Declarations
'
Public Enum eWinKeys            '## Left out by Microsoft!
    vbKeyWinLeftLogo = &H5B
    vbKeyWinRightLogo = &H5C
    vbKeyWinMenu = &H5D
End Enum

Public Enum eContextMenuMode    '## Right-Click Menu Timing
    [Before Click] = 0
    [After Click] = 1
End Enum

Public Enum eIsScrollbarVisible '## Scroll visible
    [Horizontal] = 1
    [Vertical] = 2
End Enum

Public Enum eCodeScrollView     '## Scroll Treeview
    [Home] = 0
    [Page Up] = 1
    [Up] = 2
    [Down] = 3
    [Page Down] = 4
    [End] = 5
    [Left] = 6
    [Page Left] = 7
    [Line Left] = 8
    [Line Right] = 9
    [Page Right] = 10
    [Right] = 11
End Enum

Public Enum eNodeShuffle        '## Branch level node shuffle
    [Node Up] = 0
    [Node Down] = 1
    [Node First] = 2
    [Node Last] = 3
End Enum

Public Enum eNodeDepth          '## IsChild Node Depth
    [First Branch] = 0
    [All Branches] = 1
End Enum

Public Enum eAppearance         '## Treeview Border Appearance
    [None] = 0
    [Fixed Single] = 1
    [Thin] = 2
End Enum

Public Enum eOpMode             '## Normal or Multi-Node Select
    [Single Select] = 0
    [Multi Select] = 1
End Enum

Public Enum eFileMode           '## Load/Save
    [Binary] = 0
    [xml] = 1
End Enum

Public Enum eItemOverlayIndices '## Overlay Icons. [Note: Additional Overlay modes can be used...]
  [No Overlay] = 0
  [Share Overlay] = 1
  [Shortcut Overlay] = 2
  [Custom 1] = 3
  [Custom 2] = 4
  [Custom 3] = 5
End Enum

'===========================================================================
' Private: Variables and Declarations
'
Private WithEvents oTree  As MSComctlLib.TreeView
Attribute oTree.VB_VarHelpID = -1

Private meAppearance      As eAppearance        '## TreeView Appearance
Private meContextMenuMode As eContextMenuMode   '## Before/After ContextMenu State
Private mbEnabled         As Boolean            '## TreeView Enabled state
Private mbLocked          As Boolean            '## TreeView Locked state
Private mbNodeClickEvent  As Boolean            '## Used by new TreeView Locked state to
                                                '   disable the Node Check feature.
Private mlNodeHeight      As Long               '## Physical node height (pixels)
Private mbToolTips        As Boolean            '## TreeView ToolTip Enabled/Disabled state
Private meOpMode          As eOpMode            '## Normal or Multi-Node Select state

Private moHoverNode       As MSComctlLib.Node   '## Used by Hover Event code to stop cascading events
Private moCheckNode       As MSComctlLib.Node   '## Used by new TreeView Locked state to
                                                '   disable the Node Check feature.
Private moFocusNode       As MSComctlLib.Node   '## Used in Multi-Node Select mode

Private Type NodeInfo                           '## Mimics the Node object + state info.
                                                '   Used by Binary Load/Save function.
    Index        As Long
    LinkId       As Long                        '## Parent Node Index
    Text         As String
    Key          As Variant
    Image        As Variant
    SelImage     As Variant
    ExpImage     As Variant
    Expanded     As Boolean                     '## Expanded State
    Checked      As Boolean                     '## Checked state
    Bold         As Boolean                     '## Bold state
    BackColor    As Long
    ForeColor    As Long
    Sorted       As Long
    Tag          As Variant
    Selected     As Boolean
    FirstVisible As Boolean                     '## Whether node is the first visible node
End Type

'===========================================================================
' Private: Variables and Declarations for Drag'n'Drop
'
Private WithEvents moDragScroll As XTimer
Attribute moDragScroll.VB_VarHelpID = -1
Private WithEvents moDragExpand As XTimer
Attribute moDragExpand.VB_VarHelpID = -1

Private mbDragEnabled As Boolean            '## Allow Drag'n'Drop functionality?
Private mbStartDrag   As Boolean            '## Drag operation flag
Private mbInDrag      As Boolean            '## Drag operation flag
Private mlAutoScroll  As Long               '## Distance in which auto-scrolling happens
Private mptBtnDown    As POINTAPI           '## Use to determine if MouseDown+MouseMove is
                                            '   outside start drag region.
Private mszDrag       As Size               '## X and Y distance cursor moves before
                                            '   dragging begins, in pixels.

Private moDragNode    As MSComctlLib.Node   '## Selected Drag Node

'===========================================================================
' Private: Variables and Declarations for iMultiSelect
'
Private mlSelBackColor    As OLE_COLOR      '## Custom selection Back Colour
Private mlSelForeColor    As OLE_COLOR      '## Custom selection Fore Colour
Private mbSelNoClear      As Boolean        '## SpaceBar operation flag
Private mbSelNoDefault    As Boolean        '## Focus node select state flag
Private mbSelBold         As Boolean        '## Selection bold state

Private mbSelKbdMode      As Boolean        '## Keyboard selection state flag
Private mbSelMode         As Boolean        '## Selection mode state flag
Private moSelNodes        As cTags          '## Selection Node collection
                                            '   object tracks actual node states
Private mSelectedNodes    As Collection     '## Selected node collection object

Private Enum eNodeTags                      '## moSelNodes variant fields
    [NodeKey] = 0
    [ForeColor] = 1
    [BackColor] = 2
    [Bold] = 3
End Enum

'===========================================================================
' Implemented Interfaces
'
Implements iMultiSelect

'===========================================================================
' Private: API Delclarations & variables
'
Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type POINTAPI   ' pt
  x As Long
  y As Long
End Type

Private Type Size
  cx As Long
  cy As Long
End Type

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
                                        (ByVal hWnd As Long, _
                                         ByVal msg As Long, _
                                         ByVal wParam As Long, _
                                         ByVal lParam As Long) As Long

Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" _
                                        (ByVal hWnd As Long, _
                                         ByVal wMsg As Long, _
                                         ByVal wParam As Long, _
                                               lParam As Any) As Long

Private Declare Function SendMessageAAny Lib "user32" Alias "SendMessageA" _
                                        (ByVal hWnd As Long, _
                                         ByVal wMsg As Long, _
                                               wParam As Any, _
                                               lParam As Any) As Long

Private Declare Function SetWindowPos Lib "user32" _
                                        (ByVal hWnd As Long, _
                                         ByVal hWndInsertAfter As Long, _
                                         ByVal x As Long, _
                                         ByVal y As Long, _
                                         ByVal cx As Long, _
                                         ByVal cy As Long, _
                                         ByVal wFlags As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
                                        (ByVal hWnd As Long, _
                                         ByVal nIndex As Long, _
                                         ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                                        (ByVal hWnd As Long, _
                                         ByVal nIndex As Long) As Long

Private Declare Function ImageList_SetOverlayImage Lib "comctl32.dll" _
                                        (ByVal hIml As Long, _
                                         ByVal iImage As Long, _
                                         ByVal iOverlay As Long) As Boolean

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As Any) As Long  ' lpPoint As POINTAPI) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lprc As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal fnBar As SB_Type, lpsi As SCROLLINFO) As Boolean

Private Const WM_SETREDRAW        As Long = &HB

Private Const TV_FIRST            As Long = &H1100

Private Const TVM_DELETEITEM      As Long = (TV_FIRST + 1)
Private Const TVM_GETITEMRECT     As Long = (TV_FIRST + 4)
Private Const TVM_GETNEXTITEM     As Long = (TV_FIRST + 10)
Private Const TVM_SELECTITEM      As Long = (TV_FIRST + 11)
Private Const TVM_GETITEM         As Long = (TV_FIRST + 12)
Private Const TVM_SETITEM         As Long = (TV_FIRST + 13)

' TVM_GETNEXTITEM wParam values
Private Enum TVGN_Flags
    TVGN_ROOT = &H0
    TVGN_NEXT = &H1
    TVGN_PREVIOUS = &H2
    TVGN_PARENT = &H3
    TVGN_CHILD = &H4
    TVGN_FIRSTVISIBLE = &H5
    TVGN_NEXTVISIBLE = &H6
    TVGN_PREVIOUSVISIBLE = &H7
    TVGN_DROPHILITE = &H8
    TVGN_CARET = &H9
'#If (WIN32_IE >= &H400) Then   ' >= Comctl32.dll v4.71
    TVGN_LASTVISIBLE = &HA
'#End If
End Enum

Private Type TVITEM   ' was TV_ITEM
  Mask           As Long
  hitem          As Long
  State          As Long
  StateMask      As Long
  pszText        As Long    ' pointer
  cchTextMax     As Long
  iImage         As Long
  iSelectedImage As Long
  cChildren      As Long
  lParam         As Long
End Type

Private Enum TVITEM_mask
    TVIF_TEXT = &H1
    TVIF_IMAGE = &H2
    TVIF_PARAM = &H4
    TVIF_STATE = &H8
    TVIF_HANDLE = &H10
    TVIF_SELECTEDIMAGE = &H20
    TVIF_CHILDREN = &H40
#If (WIN32_IE >= &H400) Then   ' WIN32_IE = 1024 (>= Comctl32.dll v4.71)
    TVIF_INTEGRAL = &H80
#End If
    TVIF_DI_SETITEM = &H1000   ' Notification
End Enum

Private Enum TVITEM_state
    TVIS_SELECTED = &H2
    TVIS_CUT = &H4
    TVIS_DROPHILITED = &H8
    TVIS_BOLD = &H10
    TVIS_EXPANDED = &H20
    TVIS_EXPANDEDONCE = &H40
#If (WIN32_IE >= &H300) Then
    TVIS_EXPANDPARTIAL = &H80
#End If
    
    TVIS_OVERLAYMASK = &HF00
    TVIS_STATEIMAGEMASK = &HF000
    TVIS_USERMASK = &HF000
End Enum

Private Const TVS_NOTOOLTIPS    As Long = &H80

' scrollbar definitions
Private Type SCROLLINFO
  cbSize    As Long
  fMask     As SIF_Mask
  nMin      As Long
  nMax      As Long
  nPage     As Long
  nPos      As Long
  nTrackPos As Long
End Type

Private Enum SIF_Mask
  SIF_RANGE = &H1
  SIF_PAGE = &H2
  SIF_POS = &H4
  SIF_DISABLENOSCROLL = &H8
  SIF_TRACKPOS = &H10
  SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
End Enum

Private Enum ScrollDirectionFlags
    sdLeft = &H1
    sdUp = &H2
    sdRight = &H4
    sdDown = &H8
End Enum

Private Enum SB_Type
  SB_HORZ = 0
  SB_VERT = 1
  SB_CTL = 2
  SB_BOTH = 3
End Enum

Private Const WS_HSCROLL = &H100000
Private Const WS_VSCROLL = &H200000

' Scroll Bar Commands
Private Const SB_LINEUP         As Long = 0
Private Const SB_LINELEFT       As Long = 0
Private Const SB_LINEDOWN       As Long = 1
Private Const SB_LINERIGHT      As Long = 1
Private Const SB_PAGEUP         As Long = 2
Private Const SB_PAGELEFT       As Long = 2
Private Const SB_PAGEDOWN       As Long = 3
Private Const SB_PAGERIGHT      As Long = 3
Private Const SB_THUMBPOSITION  As Long = 4
Private Const SB_THUMBTRACK     As Long = 5
Private Const SB_TOP            As Long = 6
Private Const SB_LEFT           As Long = 6
Private Const SB_BOTTOM         As Long = 7
Private Const SB_RIGHT          As Long = 7
Private Const SB_ENDSCROLL      As Long = 8

Private Const WM_HSCROLL        As Long = &H114
Private Const WM_VSCROLL        As Long = &H115

Private Const GW_CHILD          As Long = 5

'## Border Styles
Private Const GWL_EXSTYLE       As Long = (-20)
Private Const GWL_STYLE         As Long = (-16)
Private Const SWP_NOMOVE        As Long = &H2
Private Const SWP_NOSIZE        As Long = &H1
Private Const SWP_NOZORDER      As Long = &H4
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const SWP_FRAMECHANGED  As Long = &H20       '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_NOOWNERZORDER As Long = &H200      '  Don't do owner Z ordering

Private Const WS_EX_CLIENTEDGE  As Long = &H200
Private Const WS_EX_STATICEDGE  As Long = &H20000

Private Const SM_CXDRAG         As Long = &H44
Private Const SM_CYDRAG         As Long = &H45

Private Enum RectFlags
  rfLeft = &H1
  rfTop = &H2
  rfRight = &H4
  rfBottom = &H8
End Enum

'===========================================================================
' Private: cTreeView Internal Error Declarations
'
Private Const csSOURCE_ERR    As String = "cTreeView"
Private Const clTREECOPY_ERR1 As Long = vbObjectError + 2048 + 201
Private Const csTREECOPY_ERR1 As String = "Destination child node does not exist."
Private Const clTREEMODE_ERR1 As Long = vbObjectError + 2048 + 201
Private Const csTREEMODE_ERR1 As String = "MultiNode Selection disabled."
Private Const clTREEMODE_ERR2 As Long = vbObjectError + 2048 + 202
Private Const csTREEMODE_ERR2 As String = "Drag'n'Drop operations cannot be used with MultiNode Selection enabled."
Private Const clTREEHOOK_ERR  As Long = vbObjectError + 2048 + 203                              '@@ v02.01.00
Private Const csTREEHOOK_ERR  As String = "Invalid control! Must be a VB6 TreeView control"     '@@ v02.01.00

'===========================================================================
' Public: cTreeView Event Declarations
'
Public Event Selected(Node As MSComctlLib.Node)
Attribute Selected.VB_Description = "More reliable than NodeClick Event (preferred method)."
Public Event Hover(Node As MSComctlLib.Node)
Attribute Hover.VB_Description = "Generated if Treeview's HotTracking is enabled and returns the Node under the cursor."
Public Event ContextMenu(Node As MSComctlLib.Node, x As Single, y As Single)
Attribute ContextMenu.VB_Description = "Occurs either before or after the KeyPress or mouse Click events."

Public Event StartDrag(SourceNode As MSComctlLib.Node)
Attribute StartDrag.VB_Description = " Occurs when a drag-and-drop operation is initiated."
Public Event Dragging(SourceNode As MSComctlLib.Node, TargetParent As MSComctlLib.Node)
Public Event Dropped(SourceNode As MSComctlLib.Node, TargetParent As MSComctlLib.Node)
Attribute Dropped.VB_Description = "Occurs when a drag-and-drop operation is completed."

Public Event CopyNode(DestNode As MSComctlLib.Node, SrcNode As MSComctlLib.Node, Cancel As Boolean)
Attribute CopyNode.VB_Description = "Generated during the NodeCopy Function."

'===========================================================================
' Public: MSComctlLib.TreeView Event Declarations
'
Public Event AfterLabelEdit(Cancel As Integer, NewString As String)
Attribute AfterLabelEdit.VB_Description = "Occurs after a user edits the label of the currently selected Node or ListItem object."
Public Event BeforeLabelEdit(Cancel As Integer)
Attribute BeforeLabelEdit.VB_Description = "Occurs when a user attempts to edit the label of the currently selected ListItem or Node object."
Public Event Click()
Attribute Click.VB_Description = " Occurs when the user presses and then releases a mouse button over an object."
Public Event Collapse(ByVal Node As MSComctlLib.Node)
Attribute Collapse.VB_Description = "Generated when any Node object in a TreeView control is collapsed."
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when you press and release a mouse button and then press and release it again over an object."
Public Event Expand(ByVal Node As MSComctlLib.Node)
Attribute Expand.VB_Description = "Occurs when a Node object in a TreeView control is expanded; that is, when its child nodes become visible."
Public Event KeyDown(KeyCode, Shift)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Public Event KeyPress(KeyAscii)
Attribute KeyPress.VB_Description = " Occurs when the user presses and releases an ANSI key."
Public Event KeyUp(KeyCode, Shift)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Public Event MouseDown(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single, Cancel As Boolean)
Attribute MouseDown.VB_Description = " Occurs when the user presses the mouse button while an object has the focus with Optional Cancel feature."
Public Event MouseMove(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Public Event MouseUp(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Public Event NodeClick(ByVal Node As MSComctlLib.Node)
Attribute NodeClick.VB_Description = "Occurs when a Node object is clicked."
Public Event SelChange()
Attribute SelChange.VB_Description = "MultiSelect Event. Selection change notification."

'===========================================================================
' TreeView: Events
'
Private Sub oTree_AfterLabelEdit(Cancel As Integer, NewString As String)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::AfterLabelEdit -> Cancel="; CStr(Cancel); "  NewString="; NewString
    #End If
    RaiseEvent AfterLabelEdit(Cancel, NewString)
End Sub

Private Sub oTree_BeforeLabelEdit(Cancel As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::BeforeLabelEdit -> Cancel="; CStr(Cancel)
    #End If
    If mbSelMode Then
        Cancel = True
    Else
        RaiseEvent BeforeLabelEdit(Cancel)
    End If

End Sub

Private Sub oTree_Click()
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::Click"
    #End If

    '## If Locked state = true, then try and stop user from changing
    '   a Node's check state.
    If mbLocked Then
        If oTree.Checkboxes Then
            If Not mbNodeClickEvent Then
                If Not (moCheckNode Is Nothing) Then
                    moCheckNode.Checked = (Not moCheckNode.Checked)
                    Set moCheckNode = Nothing
                End If
            End If
        End If
    End If
    mbNodeClickEvent = False
    RaiseEvent Click

End Sub

Private Sub oTree_Collapse(ByVal Node As MSComctlLib.Node)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::Collapse -> Node="; Node.Text
    #End If
    RaiseEvent Collapse(Node)

End Sub

Private Sub oTree_DblClick()
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::DblClick"
    #End If
    RaiseEvent DblClick

End Sub

Private Sub oTree_DragDrop(Source As Control, x As Single, y As Single)

    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::DragDrop -> Source="; Source.Name; "  X="; CStr(x); "  Y="; CStr(y)
    #End If

    If mbDragEnabled Then
        With oTree
            .DropHighlight = .HitTest(x, y)
            If Not (.DropHighlight Is Nothing) Then                 '## Did we drop a node?
                If moDragNode <> .DropHighlight Then                '## Yes. Did we drag the node onto itself?
                    RaiseEvent Dropped(moDragNode, .DropHighlight)  '## Notify programmer & Reset
                End If
            End If
            '## Reset
            Set .DropHighlight = Nothing
            Set moDragNode = Nothing
            mbInDrag = False
            mbStartDrag = False
            '.Drag vbEndDrag                                         '!! Moved to otree_MouseUp
        End With
    End If

End Sub

Private Sub oTree_DragOver(Source As Control, x As Single, y As Single, State As Integer)

    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::DragOver -> Source="; Source.Name; "  X="; CStr(x); "  Y="; CStr(y)
    #End If

    If mbDragEnabled Then
        With oTree
            Set .DropHighlight = .HitTest(x, y)
            If .DropHighlight Is Nothing Then
                #If NODLL = 0 Then
                    .DragIcon = LoadResPicture(1, vbResCursor)
                #Else
                    .DragIcon = LoadPicture(App.Path + "\NoSelect.CUR")
                #End If
            Else
                .DragIcon = moDragNode.CreateDragImage
            End If
        End With
        pDoDrag
    End If

End Sub

Private Sub oTree_Expand(ByVal Node As MSComctlLib.Node)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::Expand -> Node="; Node.Text
    #End If
    RaiseEvent Expand(Node)

End Sub

Private Sub oTree_GotFocus()
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::GotFocus"
    #End If

End Sub

Private Sub oTree_KeyDown(KeyCode As Integer, Shift As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::KeyDown -> KeyCode="; CStr(KeyCode); "  Shift="; CStr(Shift)
    #End If

    Dim rc     As RECT
    Dim snOutX As Single
    Dim snOutY As Single

    RaiseEvent KeyDown(KeyCode, Shift)

    '{ MultiNode Select code
    If meOpMode = [Multi Select] Then
        mbSelKbdMode = True
    End If
    '}

    '## Fire ContextMenu event before click event?
    If meContextMenuMode = [Before Click] Then
        Select Case KeyCode
            Case vbKeyWinMenu
                '
                '## Retrieve selected node's co-ordinates and covert to the treeview's parent
                '   container's scale mode.
                '
                rc = GetSelectedNodeCoords
                pScale CSng(rc.Left + (rc.Right - rc.Left) \ 4), _
                       CSng(rc.Top + (rc.Bottom - rc.Top) \ 2), _
                       snOutX, _
                       snOutY, _
                       vbPixels, _
                       oTree.Parent.ScaleMode
                RaiseEvent ContextMenu(oTree.SelectedItem, snOutX, snOutY)
        End Select
    End If

End Sub

Private Sub oTree_KeyPress(KeyAscii As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::KeyPress -> KeyAscii="; CStr(KeyAscii)
    #End If

    '{ MultiNode Select code
    If meOpMode = [Multi Select] Then
        Select Case KeyAscii
            Case vbKeyEscape
                KeyAscii = 0
            Case vbKeySpace
                KeyAscii = 0
        End Select
    End If
    '}

    If KeyAscii Then
        RaiseEvent KeyPress(KeyAscii)
    End If

End Sub

Private Sub oTree_KeyUp(KeyCode As Integer, Shift As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::KeyUp -> KeyCode="; CStr(KeyCode); "  Shift="; CStr(Shift)
    #End If

    Dim rc     As RECT
    Dim snOutX As Single
    Dim snOutY As Single

    '{ MultiNode Select code
    If meOpMode = [Multi Select] Then
        Select Case KeyCode
            Case vbKeyEscape
                If Shift = vbShiftMask Then
                    iMultiSelect_ClearSelection
                    KeyCode = 0
                End If
            Case vbKeySpace
                Select Case Shift
                    Case 0
                        iMultiSelect_ToggleNode , CBool(Shift = vbShiftMask)
                        KeyCode = False
                    Case vbShiftMask
                        iMultiSelect_ToggleNode moFocusNode, CBool(Shift = vbShiftMask)
                        KeyCode = False
                End Select
            Case vbKeyA
                If Shift = vbCtrlMask Then
                    iMultiSelect_SelectAllNodes
                    KeyCode = 0
                End If
            Case vbKeyT
                If Shift = vbCtrlMask Then
                    iMultiSelect_ToggleSelection
                    KeyCode = 0
                End If
        End Select
    End If
    '}

    If KeyCode Then
        RaiseEvent KeyUp(KeyCode, Shift)
    End If

    '## Fire ContextMenu event after keypress event?
    If meContextMenuMode = [After Click] Then
        Select Case KeyCode
            Case vbKeyWinMenu
                '
                '## Retrieve selected node's co-ordinates and covert to the treeview's parent
                '   container's scale mode.
                '
                rc = GetSelectedNodeCoords
                pScale CSng(rc.Left + (rc.Right - rc.Left) \ 4), CSng(rc.Top + (rc.Bottom - rc.Top) \ 2), _
                       snOutX, snOutY, _
                       vbPixels, _
                       oTree.Parent.ScaleMode
                RaiseEvent ContextMenu(oTree.SelectedItem, snOutX, snOutY)
        End Select
    End If

End Sub

Private Sub oTree_LostFocus()
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::LostFocus"
    #End If

End Sub

Private Sub oTree_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    #If DEBUGMODE = 1 Then
        #If MOUSEEVENTS = 1 Or MOUSEEVENTS = 2 Then
            Debug.Print dbgCtrlName; "::MouseDown -> Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y)
        #End If
    #End If

    Dim oNode   As MSComctlLib.Node
    Dim bCancel As Boolean
    
    With oTree
        '## Is drag'n'drop allowed?
        If mbDragEnabled Then
            GetCursorPos mptBtnDown                         '## Update cursor position for Drag'n'Drop testing
            If Button = vbLeftButton Then
                Set moDragNode = .HitTest(x, y)             '## Capture the node to be dragged
            End If
        End If

        '## Fire ContextMenu event before click event?
        If meContextMenuMode = [Before Click] Then
            If Button = vbRightButton Then
                Set .SelectedItem = .HitTest(x, y)
                RaiseEvent ContextMenu(.SelectedItem, x, y)
            End If
        End If
    End With

    RaiseEvent MouseDown(Button, Shift, x, y, bCancel)

    '{ MultiNode Select code
    mbSelMode = False
    mbSelKbdMode = False
    If meOpMode = [Multi Select] Then
        If Not bCancel Then
            Set oNode = oTree.HitTest(x, y)
            If (Shift And vbCtrlMask) = vbCtrlMask Then
                mbSelMode = True
            ElseIf (Not mbSelNoClear) Or (Not mbSelKbdMode) Then
                iMultiSelect_ClearSelection
                RaiseEvent SelChange
            End If
        End If
    End If
    '}

End Sub

Private Sub oTree_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    #If DEBUGMODE = 1 Then
        #If MOUSEEVENTS = 2 Then
            Debug.Print dbgCtrlName; "::MouseMove -> Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y)
        #End If
    #End If

    Dim oNode As MSComctlLib.Node
    Dim pt As POINTAPI

    Set oNode = oTree.HitTest(x, y)
    Select Case Button
        Case vbLeftButton
        Case vbRightButton
        Case Else
            '
            '## Hover event
            '
            If oTree.HotTracking Then
                '## Requires HotTracking
                If Not (oNode Is Nothing) Then
                    '## Is the cursor is over a node
                    Select Case True
                        Case (moHoverNode Is Nothing), (oNode.Key <> moHoverNode.Key)
                            '## New node selected
                            Set moHoverNode = oNode
                            RaiseEvent Hover(oNode)
                    End Select
                End If
            End If

    End Select

    On Error GoTo ErrorHandler

    If mbDragEnabled Then                                   '## Is drag'n'drop allowed?
        If Button = vbLeftButton Then                       '## Yes. Signal a Drag operation.
            With oTree
                If Not (.HitTest(x, y) Is Nothing) Then     '## Do we have a node selected?
                    If mbStartDrag = True Then
                        mbInDrag = True                         '## Yes. Set the flag to true.
                        '.DragIcon = moDragNode.CreateDragImage '!! Moved to otree_DragOver
                        .Drag vbBeginDrag                       '## Signal VB to start drag operation.
                    Else
                        If Not (moDragNode Is Nothing) Then
                            GetCursorPos pt
                            If (Abs(pt.x - mptBtnDown.x) >= mszDrag.cx) Or (Abs(pt.y - mptBtnDown.y) >= mszDrag.cy) Then
                                RaiseEvent StartDrag(moDragNode)    '## Notify programmer starting drag operation
                                mbStartDrag = True
                            End If
                        End If
                    End If
                End If
            End With
        End If
    End If

    RaiseEvent MouseMove(Button, Shift, x, y)
    Exit Sub

ErrorHandler:
    mbStartDrag = False
    mbInDrag = False
End Sub

Private Sub oTree_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    #If DEBUGMODE = 1 Then
        #If MOUSEEVENTS = 1 Or MOUSEEVENTS = 2 Then
            Debug.Print dbgCtrlName; "::MouseUp -> Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y)
        #End If
    #End If
    Dim oNode As MSComctlLib.Node

    Set oNode = oTree.HitTest(x, y)
    Select Case Button
        Case vbLeftButton
            '{ MultiNode Select code
            If meOpMode = [Multi Select] Then
                If Not (oNode Is Nothing) Then
                    If (Shift And vbCtrlMask) = vbCtrlMask Then
                        iMultiSelect_ToggleNode oNode, CBool((Shift And vbShiftMask) = vbShiftMask)
                        oNode.Selected = False
                    End If
                End If
            End If
            '}

        Case vbRightButton

        Case Else
            '{ MultiNode Select code
            If mbSelMode Then
                If Not (oNode Is Nothing) Then
                    oNode.Selected = False
                End If
            End If
            '}

    End Select
    '{ MultiNode Select code
    Set moFocusNode = oNode
    '}
    '## Fire ContextMenu event after click event?
    If meContextMenuMode = [After Click] Then
        If Button = vbRightButton Then
            With oTree
                Set .SelectedItem = .HitTest(x, y)
                RaiseEvent ContextMenu(.SelectedItem, x, y)
            End With
        End If
    End If

    If mbDragEnabled Then                                   '## Is drag'n'drop allowed?
        mbStartDrag = False
        mbInDrag = False

        oTree.Drag vbEndDrag
        moDragExpand.Enabled = False
        moDragScroll.Enabled = False
    End If

    RaiseEvent MouseUp(Button, Shift, x, y)

End Sub

Private Sub oTree_NodeCheck(ByVal Node As MSComctlLib.Node)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::NodeCheck -> Node="; Node.Text
    #End If

    Set moCheckNode = Node

End Sub

Private Sub oTree_NodeClick(ByVal Node As MSComctlLib.Node)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::NodeClick -> Node="; Node.Text
    #End If
    RaiseEvent NodeClick(Node)

    '## Node selected, raise event
    mbNodeClickEvent = True

    '{ MultiNode Select code
    If meOpMode = [Multi Select] Then
        If Not mbSelMode And Not mbSelKbdMode Then
            If mbSelNoDefault Then
                Node.Selected = False
            End If
            With Node
                '## Reset collections and select node
                Set moSelNodes = New cTags
                Set mSelectedNodes = New Collection
                pSetNodeSelect Node, True
            End With
            RaiseEvent SelChange
        End If
        If mbSelKbdMode And mbSelNoDefault Then
            Node.Selected = False
        End If
        Set moFocusNode = Node
    End If
    '}

    RaiseEvent NodeClick(Node)
    RaiseEvent Selected(Node)

End Sub

Private Sub oTree_OLECompleteDrag(Effect As Long)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLECompleteDrag -> Effect="; CStr(Effect)
    #End If

End Sub

Private Sub oTree_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLEDragDrop -> Effect="; CStr(Effect); "  Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y)
    #End If

End Sub

Private Sub oTree_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLEDragOver -> Effect="; CStr(Effect); "  Button="; CStr(Button); "  Shift="; CStr(Shift); "  X="; CStr(x); "  Y="; CStr(y); "  State="; CStr(State)
    #End If

End Sub

Private Sub oTree_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLEGiveFeedback -> Effect="; CStr(Effect); "  DefaultCursors="; CStr(DefaultCursors)
    #End If

End Sub

Private Sub oTree_OLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLESetData -> Effect="; CStr(DataFormat)
    #End If

End Sub

Private Sub oTree_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::OLEStartDrag -> AllowedEffects="; CStr(AllowedEffects)
    #End If

End Sub

Private Sub oTree_Validate(Cancel As Boolean)
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::Validate -> Cancel="; CStr(Cancel)
    #End If

End Sub

'===========================================================================
' Internal Class Subroutines
'
Private Sub Class_Initialize()
    '## Set defaults...
    meContextMenuMode = [After Click]
    meOpMode = [Single Select]          '## Normal TreeView operation mode
    mbToolTips = True

    '{ MultiNode Select code
    mbSelNoDefault = True
    mlSelBackColor = vbHighlight
    mlSelForeColor = vbHighlightText
    Set moSelNodes = New cTags
    Set mSelectedNodes = New Collection
    '}

End Sub

Private Sub Class_Terminate()
    '## release control
    UnHookCtrl
    '{ MultiNode Select code
    Set moSelNodes = Nothing
    Set mSelectedNodes = Nothing
    '}
    If Not (moDragScroll Is Nothing) Then Set moDragScroll = Nothing
    If Not (moDragExpand Is Nothing) Then Set moDragExpand = Nothing
End Sub

'===========================================================================
' Timers: Events
'
Private Sub moDragExpand_Tick()
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::moDragExpand -> Tick!"
    #End If
    With oTree
        Select Case True
            Case (.DropHighlight Is Nothing), (moDragNode Is Nothing)
                '## Avoid possible error - should not get here!
            Case (.DropHighlight.Children > 0) And (.DropHighlight.Expanded = False)
                '## The highlighted node has children and is collapsed
                .DropHighlight.Expanded = True
        End Select
    End With
    '## Kill Timer
    moDragExpand.Enabled = False
End Sub

Private Sub moDragScroll_Tick()
    #If DEBUGMODE = 1 Then
        Debug.Print dbgCtrlName; "::moDragScroll -> Tick!"
    #End If

    Dim pt            As POINTAPI
    Dim rcClient      As RECT
    Dim dwRectFlags   As RectFlags
    Dim dwScrollFlags As ScrollDirectionFlags

    If mbInDrag = False Then
        '## We shouldn't be here!
        moDragScroll.Enabled = False
        Exit Sub
    End If
    '
    '## Get the cursor postion in TreeView client coords
    '
    With oTree
        GetCursorPos pt
        ScreenToClient .hWnd, pt
        GetClientRect .hWnd, rcClient
    End With
    '
    '## If the cursor is within an auto scroll region in the TreeView's client area...
    '
    dwRectFlags = pPtInRectRegion(rcClient, mlAutoScroll, pt)
    If dwRectFlags Then
        '
        '## Determine which direction the TreeView can be scrolled...
        '
        dwScrollFlags = pIsWindowScrollable(oTree.hWnd)
        '
        '## If the cursor is within the respective drag region specified by the
        '   mlAutoScroll distance, and if the TreeView can be scrolled
        '   in that direction, send the TreeView that respective scroll message.
        '
        Select Case True
            Case (dwRectFlags And rfLeft) And (dwScrollFlags And sdLeft)
                Debug.Print "Left"
                ScrollView [Line Left]
            Case (dwRectFlags And rfRight) And (dwScrollFlags And sdRight)
                Debug.Print "Right"
                ScrollView [Line Right]
            Case (dwRectFlags And rfTop) And (dwScrollFlags And sdUp)
                Debug.Print "Up"
                ScrollView [Up]
            Case (dwRectFlags And rfBottom) And (dwScrollFlags And sdDown)
                Debug.Print "Down"
                ScrollView [Down]
            Case Else
                moDragScroll.Enabled = False
        End Select
    End If

End Sub

'===========================================================================
' Public Properties
'
Public Property Get Appearance() As eAppearance
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D, Flat 3D  or no effects."
    Appearance = meAppearance
End Property

Public Property Let Appearance(ByVal New_Appearance As eAppearance)

    Dim lStyle As Long

    meAppearance = New_Appearance

    '## Retrieve current style
    lStyle = GetWindowLong(oTree.hWnd, GWL_EXSTYLE)

    '## Toggle style flags
    Select Case New_Appearance
        Case [None]:         lStyle = lStyle And Not WS_EX_CLIENTEDGE And Not WS_EX_STATICEDGE
        Case [Fixed Single]: lStyle = lStyle Or WS_EX_CLIENTEDGE And Not WS_EX_STATICEDGE
        Case [Thin]:         lStyle = lStyle And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    End Select

    '## Apply changes
    SetWindowLong oTree.hWnd, GWL_EXSTYLE, lStyle
    SetWindowPos oTree.hWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or _
                                            SWP_NOSIZE Or SWP_NOMOVE

End Property

Public Property Get ChildParentKeys(Optional ByVal oNode As MSComctlLib.Node) As String()
Attribute ChildParentKeys.VB_Description = "Returns a string array of node keys."

    Dim skeys() As String
    Dim sTmp()  As String
    Dim iMax    As Integer
    Dim iLoop   As Integer

    '## Check and set Node
    If (oNode Is Nothing) Then
        If (oTree.SelectedItem Is Nothing) Then
            Exit Property
        Else
            Set oNode = oTree.SelectedItem
        End If
    Else
        Set oNode = oNode
    End If

    '## Traverse backwards and store keys
    Do
        iMax = iMax + 1
        ReDim Preserve sTmp(1 To iMax)
        sTmp(iMax) = oNode.Key
        If Not (oNode.FullPath = oNode.Text) Then
            Set oNode = oNode.Parent
        Else
            '## Root node reached.
            '   Swap keys to correct sequence and return results
            ReDim Preserve skeys(1 To iMax)
            For iLoop = 1 To iMax
                skeys(iLoop) = sTmp(iMax - (iLoop - 1))
            Next
            ChildParentKeys = skeys()
            Exit Do
        End If
    Loop

End Property

Public Property Get ContextMenuMode() As eContextMenuMode
Attribute ContextMenuMode.VB_Description = "Set whether the ContextMenu Event fires before or after KeyPress or Click Events [Default = After]."
    ContextMenuMode = meContextMenuMode
End Property

Public Property Let ContextMenuMode(ByVal Mode As eContextMenuMode)
    meContextMenuMode = Mode
End Property

#If NODLL = 0 Then
    Public Property Get Ctrl() As Object
Attribute Ctrl.VB_Description = "Exposes the TreeView's properties and methods that are not wrapped by the cTreeView class"
        '## Default Object
        Set Ctrl = oTree
#Else
    Public Property Get Ctrl() As MSComctlLib.TreeView
        '## Default Object
        Set Ctrl = oTree
#End If

End Property

Public Property Get CutIcon(Optional ByVal Node As MSComctlLib.Node) As Boolean
Attribute CutIcon.VB_Description = "Gets/Sets a Node's cut icon state."

    Dim tvi   As TVITEM
    Dim hitem As Long

    '## Get Node handle
    hitem = pGetNodeHandle(Node)
    '## Process if valid handle
    If hitem Then
        '## Retrieve Cut Icon Flag
        tvi.Mask = TVIF_STATE
        tvi.hitem = hitem
        tvi.StateMask = TVIS_CUT
        If SendMessageAny(oTree.hWnd, TVM_GETITEM, 0&, tvi) Then
            CutIcon = (tvi.State And TVIS_CUT) / (2 ^ 8)
        End If
    End If

End Property

Public Property Let CutIcon(Optional ByVal Node As MSComctlLib.Node, ByVal SetState As Boolean)

    Dim tvi   As TVITEM
    Dim hitem As Long
  
    '## Get Node handle
    hitem = pGetNodeHandle(Node)
    '## Process if valid handle
    If hitem Then
        With tvi
            .hitem = hitem
            .Mask = TVIF_STATE
            .State = SetState And TVIS_CUT
            '## Indicate what state bits we're changing
            .StateMask = TVIS_CUT
        End With
        '## Old docs say returns 0 on success, -1 on failure. New docs say
        '   returns TRUE if successful, or FALSE otherwise we'll go new...
        SendMessageAny oTree.hWnd, TVM_SETITEM, 0&, tvi
    End If

End Property

Public Property Get DragEnabled() As Boolean
Attribute DragEnabled.VB_Description = "Enables/disables Drag'n'Drop operation. [Note: Cannot be used with MultiSelect feature]"
    DragEnabled = mbDragEnabled
End Property

Public Property Let DragEnabled(ByVal Enabled As Boolean)
    mbDragEnabled = Enabled
End Property

Public Property Get DragExpandTime() As Long
Attribute DragExpandTime.VB_Description = "Auto Expand Parent node timer (milliseconds) during Drag'n'Drop operation [Default = 1000ms (1 seconds)]"
    DragExpandTime = moDragExpand.Interval
End Property

Public Property Let DragExpandTime(ByVal New_DragExpandTime As Long)
    moDragExpand.Interval = New_DragExpandTime
End Property

Public Property Get DragScrollTime() As Long
Attribute DragScrollTime.VB_Description = "Auto scroll timer (milliseconds) during Drag'n'Drop operation [Default = 200ms (0.2 seconds)]"
    DragScrollTime = moDragScroll.Interval
End Property

Public Property Let DragScrollTime(ByVal New_DragScrollTime As Long)
    moDragScroll.Interval = New_DragScrollTime
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "True disabled appearance! Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = mbEnabled
End Property

Public Property Let Enabled(ByVal bState As Boolean)

    Dim tvi   As TVITEM
    Dim hitem As Long

    If Not mbLocked Then
        tvi.Mask = TVIF_STATE
        hitem = SendMessageLong(oTree.hWnd, TVM_GETNEXTITEM, TVGN_ROOT, 0&)
        Do While hitem <> 0
            With tvi
                .hitem = hitem
                .Mask = TVIF_STATE
                .StateMask = TVIS_CUT
                If (bState) Then                            '## Toggle the cut icon states
                    .State = .StateMask And Not TVIS_CUT
                Else
                    .State = .StateMask Or TVIS_CUT
                End If
                SendMessageAny oTree.hWnd, TVM_SETITEM, 0&, tvi
            End With
            hitem = SendMessageLong(oTree.hWnd, TVM_GETNEXTITEM, TVGN_NEXTVISIBLE, hitem)
        Loop
    End If
    oTree.Enabled = bState
    mbEnabled = bState
End Property

Public Property Get FirstVisibleNode() As MSComctlLib.Node
Attribute FirstVisibleNode.VB_Description = "Returns the first viewable node"
    '
    '## Returns the first viewable node
    '
    Dim hitem   As Long
    Dim selNode As MSComctlLib.Node
    
    With oTree
        '## Remember the node currently selected
        Set selNode = .SelectedItem
        '## Get the handle of the first visible Node
        hitem = SendMessageAny(.hWnd, TVM_GETNEXTITEM, TVGN_FIRSTVISIBLE, ByVal 0&)
        '## Make it the selected Node
        SendMessageAny .hWnd, TVM_SELECTITEM, TVGN_CARET, ByVal hitem
        '## Return the result as a Node object
        Set FirstVisibleNode = .SelectedItem
        '## Restore node that was selected
        Set .SelectedItem = selNode
    End With

End Property

Public Property Set FirstVisibleNode(ByVal Node As MSComctlLib.Node)
    '
    '## Set the first visible Node of a TreeView control
    '
    Dim hitem   As Long
    Dim selNode As MSComctlLib.Node
    
    With oTree
        '## Remember the node currently selected
        Set selNode = .SelectedItem
        '## Make the Node the select Node in the control
        Set .SelectedItem = Node
        '## Now we can get its handle
        hitem = SendMessageAny(.hWnd, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&)
        '## Restore node that was selected
        Set .SelectedItem = selNode
        '## Make it the fist visible Node
        SendMessageAny .hWnd, TVM_SELECTITEM, TVGN_FIRSTVISIBLE, ByVal hitem
    End With

End Property

Public Property Get FullKeyPath(Optional ByVal oNode As MSComctlLib.Node) As String
Attribute FullKeyPath.VB_Description = "Returns a path full of keys instead of node.text seperated with the TreeView's Seperator."
    '
    '## Returns a node's path of keys
    '
    Dim skeys() As String
    Dim iMax    As Integer
    Dim iLoop   As Integer

    '## Retrieve keys for node passed
    skeys = ChildParentKeys(oNode)
    '## Build path string
    iMax = UBound(skeys)
    For iLoop = 1 To iMax
        FullKeyPath = FullKeyPath + skeys(iLoop) + oTree.PathSeparator
    Next
    FullKeyPath = VBA.Left$(FullKeyPath, Len(FullKeyPath) - 1)

End Property

Public Property Get LastVisibleNode() As MSComctlLib.Node
Attribute LastVisibleNode.VB_Description = "Returns the last visible node."
    '
    '## Returns the last visible node
    '
    Dim lLoop    As Long
    Dim hitem    As Long
    Dim hItemOld As Long
    Dim selNode  As MSComctlLib.Node
    
    With oTree
        '## Remember the node currently selected
        Set selNode = .SelectedItem
        '## Get the handle of the first visible Node
        hitem = SendMessageAny(.hWnd, TVM_GETNEXTITEM, TVGN_FIRSTVISIBLE, ByVal 0&)
        For lLoop = 1 To .GetVisibleCount - 1
            '## Get the handle of the next visible Node
            hitem = SendMessageAny(.hWnd, TVM_GETNEXTITEM, TVGN_NEXTVISIBLE, ByVal hitem)
            '## Was a Node handle returned?
            If hitem Then
                hItemOld = hitem
            Else
                '## Less items visible than .GetVisibleCount
                hitem = hItemOld
                Exit For
            End If
        Next
        '## Make it the selected Node
        SendMessageAny .hWnd, TVM_SELECTITEM, TVGN_CARET, ByVal hitem
        '## Return the result as a Node object
        Set LastVisibleNode = .SelectedItem
        '## Restore node that was selected
        Set .SelectedItem = selNode
    End With

End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = " Determines whether a control can be edited."
   Locked = mbLocked
End Property

Public Property Let Locked(ByVal bState As Boolean)

    Dim tvi   As TVITEM
    Dim hitem As Long
   
    mbLocked = bState

    If oTree.Enabled Then
        tvi.Mask = TVIF_STATE
        hitem = SendMessageLong(oTree.hWnd, TVM_GETNEXTITEM, TVGN_ROOT, 0&)
        Do While hitem <> 0
            With tvi
                .hitem = hitem
                .Mask = TVIF_STATE
                .StateMask = TVIS_CUT
                If (bState) Then                            '## Toggle the cut icon states
                    .State = .StateMask Or TVIS_CUT
                Else
                    .State = .StateMask And Not TVIS_CUT
                End If
                SendMessageAny oTree.hWnd, TVM_SETITEM, 0&, tvi
            End With
            hitem = SendMessageLong(oTree.hWnd, TVM_GETNEXTITEM, TVGN_NEXTVISIBLE, hitem)
        Loop
    End If

End Property

Public Property Get NodeNestingLevel(ByVal Node As MSComctlLib.Node) As Integer
Attribute NodeNestingLevel.VB_Description = "Returns number of nesting/branch levels for a specific node."
    '
    '## Returns the nesting level of a TreeView's Node object
    '   (returns zero for root nodes.)
    '
    Do Until (Node.Parent Is Nothing)
        NodeNestingLevel = NodeNestingLevel + 1
        Set Node = Node.Parent
    Loop

End Property

Public Property Get OverlayIcon(Optional ByVal Node As MSComctlLib.Node) As eItemOverlayIndices
Attribute OverlayIcon.VB_Description = "Returns/Sets a Node's Overlay Icon."

    Dim tvi   As TVITEM
    Dim hitem As Long

    '## Get Node handle
    hitem = pGetNodeHandle(Node)
    With oTree
        If hitem Then
            '## Retrieve Overlay Icon Flag
            tvi.Mask = TVIF_STATE
            tvi.hitem = hitem
            tvi.StateMask = TVIS_OVERLAYMASK
            If SendMessageAny(.hWnd, TVM_GETITEM, 0&, tvi) Then
                OverlayIcon = CByte((tvi.State And TVIS_OVERLAYMASK) / (2 ^ 8))
            End If
        End If
    End With

End Property

Public Property Let OverlayIcon(Optional ByVal Node As MSComctlLib.Node, Index As eItemOverlayIndices)

    Dim tvi   As TVITEM
    Dim hitem As Long

    '## Get Node handle
    hitem = pGetNodeHandle(Node)
    With oTree
        If hitem Then
            '## Set Overlay Icon Flag
            tvi.Mask = TVIF_STATE
            tvi.hitem = hitem
            tvi.State = CLng(Index) * (2 ^ 8)
            tvi.StateMask = TVIS_OVERLAYMASK
            SendMessageAny .hWnd, TVM_SETITEM, 0&, tvi
        End If
    End With

End Property

Public Property Get ParentNodeText(ByVal Node As MSComctlLib.Node) As String
Attribute ParentNodeText.VB_Description = "Returns the parent node text (Safe routine to stop error condition if no parent node)."
    '
    '## Returns the parent text
    '   NOTE: If a Node has no parent then VB raises an error.
    '         This is a much safer Node.Parent.Text routine.
    With Node
        ParentNodeText = IIf(.FullPath = .Text, "", .Parent.Text)
    End With
End Property

Public Property Get ToolTips() As Boolean
Attribute ToolTips.VB_Description = "Enable/Disable Treeview ToolTips."
    ToolTips = mbToolTips
End Property

Public Property Let ToolTips(ByVal NewState As Boolean)
    mbToolTips = NewState
    With oTree
        If NewState Then
            '## Enable ToolTips (Default )
            SetWindowLong .hWnd, GWL_STYLE, GetWindowLong(.hWnd, GWL_STYLE) And Not TVS_NOTOOLTIPS
        Else
            '## Disable ToolTips
            SetWindowLong .hWnd, GWL_STYLE, GetWindowLong(.hWnd, GWL_STYLE) Or TVS_NOTOOLTIPS
        End If
    End With
End Property

'===========================================================================
' Public subroutines and functions
'
Public Sub ClearTreeView()
Attribute ClearTreeView.VB_Description = "A very fast way of clearing all nodes."
    '
    '## Clear the treeview
    '
    Dim hitem As Long
    Dim hWnd  As Long

    oTree.Visible = False
    hWnd = oTree.hWnd
    Do
        hitem = SendMessageLong(hWnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
        If hitem <= 0 Then Exit Do
        SendMessageLong hWnd, TVM_DELETEITEM, 0&, hitem
    Loop
    oTree.Visible = True

End Sub

Public Sub CollapseAll()
Attribute CollapseAll.VB_Description = "Expands all parent and sub-parent nodes of the entire TreeView."
    '
    '## Collapse all nodes
    '
    Dim Node As MSComctlLib.Node

    For Each Node In oTree.Nodes
        With Node
            If .Children Then .Expanded = False
        End With
    Next

End Sub

Public Sub CollapseChildNodes(ByVal Node As MSComctlLib.Node)
Attribute CollapseChildNodes.VB_Description = "Cascading effect. Hides all children for each parent node sub-branch."

    Dim lCount As Long
    Dim lLoop  As Long

    With Node
        If Not Node Is Nothing Then
            .Expanded = False
        End If

        If .Children Then               '## Does node have children?
            lCount = .Children
            Set Node = .Child
            For lLoop = 1 To lCount
                CollapseChildNodes Node '## Yes, collapse them as well (recursive)
                If lLoop < lCount Then
                    Set Node = .Next
                End If
            Next
        End If
    End With

End Sub

Public Sub ExpandAll()
Attribute ExpandAll.VB_Description = "Expands all parent and sub-parent nodes of the entire TreeView."
    '
    '## Expand all nodes
    '
    Dim Node As MSComctlLib.Node

    On Error GoTo ErrorHandler                      '@@ v01.00.01
    For Each Node In oTree.Nodes
        With Node
            If .Children Then .Expanded = True
        End With
    Next
    Exit Sub

ErrorHandler:
    Select Case Err.Number
        Case 35606: ExpandAll
    End Select
End Sub

Public Sub ExpandChildNodes(ByVal Node As MSComctlLib.Node)
Attribute ExpandChildNodes.VB_Description = "Cascading effect. Expand all parent & sub-parent nodes that belong to Node."

    Dim lCount As Long
    Dim lLoop  As Long

    With Node
        If Not Node Is Nothing Then
            .Expanded = True
        End If

        If .Children Then               '## Does node have children?
            lCount = .Children
            Set Node = .Child
            For lLoop = 1 To lCount
                ExpandChildNodes Node   '## Yes, expand them as well (recursive)
                If lLoop < lCount Then
                    Set Node = .Next
                End If
            Next
        End If
    End With

End Sub

Private Function GetSelectedNodeCoords() As RECT

    Dim hNode As Long
    Dim rcBox As RECT

    With oTree
        If Not .SelectedItem Is Nothing Then
            '## Node must be fully visible to get coords.
            If .SelectedItem.Visible Then

                '## Get a handle to the currently selected node
                hNode = SendMessageAny(.hWnd, TVM_GETNEXTITEM, TVGN_CARET, ByVal TVGN_ROOT)

                '## Set the first four bytes (Long) of the Rect to the handle of the node, before
                '   calling TVM_GETITEMRECT.
                rcBox.Left = hNode

                '## Get the coordinates of the node
                If SendMessageAny(oTree.hWnd, TVM_GETITEMRECT, True, rcBox) = 0 Then
                    '## Zero the rect to show failure
                    With rcBox
                        .Left = 0
                        .Top = 0
                        .Right = 0
                        .Bottom = 0
                    End With
                End If
            End If
        End If
    End With

    '## Return the rect
    GetSelectedNodeCoords = rcBox

End Function

#If NODLL = 0 Then
    Public Sub HookCtrl(ByRef Ctrl As Object, Optional OpMode As eOpMode = [Single Select])
Attribute HookCtrl.VB_Description = "Primes the cTreeview class to manage the  Treeview control"

    If Not (TypeOf Ctrl Is MSComctlLib.TreeView) Then
        Err.Raise clTREEHOOK_ERR, csSOURCE_ERR, csTREEHOOK_ERR  '@@ v02.01.00
        Exit Sub
    End If

#Else
    
    Public Sub HookCtrl(ByRef Ctrl As MSComctlLib.TreeView, Optional OpMode As eOpMode = [Single Select])

#End If

    Dim bState As Boolean
    Dim rc     As RECT

    Set oTree = Ctrl                                '## Capture the Treeview control.
                                                    '   Class won't work if this isn't called first.
    With oTree
        #If DEBUGMODE = 1 Then
            dbgCtrlName = .Parent.Name + "." + .Name
        #End If
        '
        '## calculate the height (pixels) of a node
        '
        bState = .Scroll                            '## Remember if scrollbars are enabled
        .Scroll = False                             '## Turn them off
        mlNodeHeight = .Height \ .GetVisibleCount   '## calculate height
        .Scroll = bState                            '## restore scrollbar state
    End With
    meOpMode = OpMode

    Dim lStyle As Long
    lStyle = GetWindowLong(oTree.hWnd, GWL_EXSTYLE)
    '
    '## Drag'n'Drop primer
    '
    mbDragEnabled = False
    '
    '## Establish the distance in which auto-scrolling happens within
    '   the TreeView's client area (we need a root item for these calls)
    '
    rc.Left = SendMessageAny(oTree.hWnd, TVM_GETNEXTITEM, ByVal TVGN_ROOT, ByVal 0&)
    If SendMessageAny(oTree.hWnd, TVM_GETITEMRECT, ByVal 1, rc) Then
        mlAutoScroll = (rc.Bottom - rc.Top) * 2
    Else
        mlAutoScroll = 32
    End If
    '
    '## Get system Start Drag Zone metrics
    '
    mszDrag.cx = GetSystemMetrics(SM_CXDRAG)
    mszDrag.cy = GetSystemMetrics(SM_CYDRAG)
    '
    '## Initialise timers & their default timings
    '
    Set moDragScroll = New XTimer
    moDragScroll.Interval = 200     '## 0.2 seconds
    Set moDragExpand = New XTimer
    moDragExpand.Interval = 1000    '## 1 second

End Sub

Public Function IsChildNode(ByVal TestNode As MSComctlLib.Node, _
                            ByVal ParentNode As MSComctlLib.Node, _
                   Optional ByVal NodeDepth As eNodeDepth = [First Branch]) As Boolean
Attribute IsChildNode.VB_Description = "Returns if a node is a direct child of another node or whether a node exists on another node's branch."

    If NodeDepth = [First Branch] Then
        If ParentNode.Children Then
            If Not IsRootNode(TestNode) Then
                IsChildNode = (TestNode.Parent = ParentNode)
            End If
        End If
    Else
        IsChildNode = IsParentNode(TestNode, ParentNode)
    End If

End Function

Public Function IsParentNode(ByVal ChildNode As MSComctlLib.Node, _
                             ByVal ParentNode As MSComctlLib.Node) As Boolean
Attribute IsParentNode.VB_Description = "Checks if one node is the parent of another."
    '
    '## Checks if one node is the parent of another.
    '   This is a recursive routine that steps down through
    '   the branches of the parent node.
    '
    Dim lNDX As Long

    If (ParentNode Is Nothing) Or (ChildNode Is Nothing) Then
        Exit Function
    End If
    If ParentNode.Children Then             '## Does the parent node have children?
        lNDX = ParentNode.Child.Index       '## Yes, remember the first child
        Do                                  '## Step through all child nodes
            If lNDX = ChildNode.Index Then  '## is ChildNode the test node?
                IsParentNode = True         '## ParentNode is the parent of ChildNode.
                Exit Do
            End If
            If IsParentNode(ChildNode, oTree.Nodes(lNDX)) Then  '## Step down through the branches
                IsParentNode = True         '## ParentNode is the parent of ChildNode.
                Exit Do
            End If
            If lNDX <> ParentNode.Child.LastSibling.Index Then  '## Have we tested the last child node?
                lNDX = oTree.Nodes(lNDX).Next.Index             '## No. Point to the next child node
            Else
                Exit Do                                         '## Yes.
            End If
        Loop
    End If

End Function

Public Function IsRootNode(ByVal Node As MSComctlLib.Node) As Boolean
Attribute IsRootNode.VB_Description = "Check is selected node is the root node."
    '
    '## Check if selected node is a root node.
    '
    With Node
        IsRootNode = (.FullPath = .Text)
    End With

End Function

Public Function IsScrollbarVisible(WhichScroll As eIsScrollbarVisible) As Boolean

    IsScrollbarVisible = (CByte(Abs((GetWindowLong(oTree.hWnd, GWL_STYLE) And WS_HSCROLL) <> 0) * 1) + _
                               (Abs((GetWindowLong(oTree.hWnd, GWL_STYLE) And WS_VSCROLL) <> 0) * 2) _
                          And WhichScroll) = WhichScroll

End Function

Public Function Load(ByVal FileName As String, _
            Optional ByVal FileMode As eFileMode) As Boolean
Attribute Load.VB_Description = "Load Treeview from Binary or XML stored file."
    Select Case FileMode
        Case [Binary]: Load = pLoadBIN(FileName)
        Case [xml]:    Load = pLoadXML(FileName)
    End Select
End Function

Public Function NodeAdd(Optional ByVal vRelative As Variant, _
                        Optional ByVal eRelationship As MSComctlLib.TreeRelationshipConstants, _
                        Optional ByVal sKey As String, _
                        Optional ByVal sText As String, _
                        Optional ByVal vImage As Variant, _
                        Optional ByVal vSelectedImage As Variant, _
                        Optional ByVal vTag As Variant, _
                        Optional ByVal bBold As Boolean = False, _
                        Optional ByVal bChecked As Boolean = False, _
                        Optional ByVal bEnsureVisible As Boolean = False, _
                        Optional ByVal bExpanded As Boolean = True, _
                        Optional ByVal bSelected As Boolean = False, _
                        Optional ByVal bSorted As Boolean = False, _
                        Optional ByVal lForeColor As OLE_COLOR, _
                        Optional ByVal lBackColor As OLE_COLOR, _
                        Optional ByVal vExpandedImage As Variant) As MSComctlLib.Node
Attribute NodeAdd.VB_Description = "Adds node to TreeView and sets properties in one call."

    Dim Node As MSComctlLib.Node

    On Error GoTo ErrorHandler

    Set Node = oTree.Nodes.Add(vRelative, eRelationship, sKey, sText, vImage, vSelectedImage)
    With Node
        #If DEBUGMODE = 1 Then
            Debug.Print "ADD-> Text = "; .Text; "   Index = "; CStr(.Index)
        #End If
        .Tag = vTag
        .Bold = bBold
        .Checked = bChecked
        If bEnsureVisible Then .EnsureVisible
        .Expanded = bExpanded
        .Selected = bSelected
        .Sorted = bSorted
        If lForeColor = 0 Then
            .ForeColor = vbWindowText
        Else
            .ForeColor = lForeColor
        End If
        If lBackColor = 0 Then
            .BackColor = vbWindowBackground
        Else
            .BackColor = lBackColor
        End If
        If Not IsMissing(vExpandedImage) Then
            .ExpandedImage = vExpandedImage
        End If
    End With
    
    Set NodeAdd = Node
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function NodeCopy(ByVal DestNode As MSComctlLib.Node, _
                         ByVal SrcNode As MSComctlLib.Node, _
                Optional ByVal IncludeChildren As Boolean = True) As Boolean    '@@ v01.00.03
Attribute NodeCopy.VB_Description = "Framework to copy one node to another. Generates CopyNode event for user controlled copying."

    Dim oSNode  As MSComctlLib.Node
    Dim oDNode  As MSComctlLib.Node
    Dim bCancel As Boolean
    '
    '## Ensure that we don't overwrite the original nodes values
    '
    Set oSNode = SrcNode
    Set oDNode = DestNode

    If oSNode = oDNode Then
        '## Same node - therefore no point trying to copy
        Exit Function
    End If
    '
    '## Force the Nodes to be expanded before the copy operation is performed. This
    '   will allow child nodes to be loaded if done on demand
    '
    oDNode.Expanded = True
    oSNode.Expanded = True
    '
    '## Let the actual process happen externally to cTREEVIEW - a must if working with a database
    '   or multiple types of data.
    '
    RaiseEvent CopyNode(oDNode, oSNode, bCancel)     '## Get the calling routine to copy the node.
    If bCancel Then                                  '## Was the process cancelled?
        NodeCopy = False
        Exit Function                                '## Yes.
    End If
    If IncludeChildren Then                                     '## Include all child nodes?
        If oSNode.Children Then                                 '## Yes. Are there child nodes?
            Set oSNode = oSNode.Child.FirstSibling              '## Yes. Select the first child.
            Do
                If (oDNode.Child Is Nothing) Then
                    NodeCopy = False
                    Err.Raise clTREECOPY_ERR1, "COPY NODE::" + csSOURCE_ERR, csTREECOPY_ERR1
                    Exit Function
                End If
                If NodeCopy(oDNode.Child.LastSibling, oSNode, IncludeChildren) Then '## Copy child node
                    If Not (oSNode = oSNode.LastSibling) Then   '## Are there more child nodes?
                        Set oSNode = oSNode.Next                '## Yes.
                    Else
                        Exit Do
                    End If
                Else
                    NodeCopy = False
                    Exit Function
                End If
            Loop
        End If
    End If
    NodeCopy = True                                         '## Success!

End Function

Public Function NodeDelete(ByVal Node As MSComctlLib.Node, _
                  Optional ByVal bSelect As Boolean = True, _
                  Optional ByVal bIncludeChilden = True) As Boolean

    Dim lCount As Long
    Dim lLoop  As Long

    With Node
        If Not IsRootNode(Node) Then    '## Is this a root node?
            If Not bIncludeChilden Then
                '## Move the node's children to the Node's parent
                lCount = .Children
                For lLoop = 1 To lCount
                    Set oTree.Nodes(.Child.Index).Parent = .Parent
                Next
            End If
            With .Parent                '## No.
                .EnsureVisible
                .Selected = bSelect
            End With
        End If
        '## Delete node
        oTree.Nodes.Remove .Index       '## Delete the node
    End With
    NodeDelete = True

End Function

Function NodeExist(ByVal Key As String) As Boolean
Attribute NodeExist.VB_Description = "Returns whether a specified node exists."

    Dim sTmp As String

    On Local Error GoTo ErrorHandler

    sTmp = TypeName(oTree.Nodes(Key))
    NodeExist = True
    Exit Function

ErrorHandler:
'
End Function

Public Function NodeFind(ByRef oResultNode As MSComctlLib.Node, _
                         ByVal NodeText As String, _
                Optional ByVal NodeKey As String, _
                Optional ByVal bSelect As Boolean = True) As Boolean
Attribute NodeFind.VB_Description = "Finds a Node and optionally selects & ensures visibility."

    Dim Node      As MSComctlLib.Node
    Dim bCheckKey As Boolean
    Dim bKeyOK    As Boolean

    bCheckKey = (Len(Trim$(NodeKey)) > 0)
    bKeyOK = (Not bCheckKey)
    For Each Node In oTree.Nodes
        With Node
            If UCase$(.Text) = UCase$(NodeText) Then
                If bCheckKey Then bKeyOK = (UCase$(.Key) = UCase$(NodeKey))
                If bKeyOK Then
                    If bSelect Then
                        .Selected = True
                        .EnsureVisible
                    End If
                    Set oResultNode = Node
                    NodeFind = True
                    Exit For
                End If
            End If
        End With
    Next

End Function

Public Function pNodeSwap(ByVal ParentNode As MSComctlLib.Node, _
                         ByVal ChildNode As MSComctlLib.Node, _
                Optional ByVal bSelect As Boolean = True) As Boolean

    Dim lNDX   As Long
    Dim lCount As Long
    Dim lLoop  As Long

    ''Debug.Print "<<>> NODE SWAP <<>>"
    oTree.Visible = False
    With ChildNode
        If IsParentNode(ParentNode, ChildNode) Then '## Are we moving a parent node?
            If IsRootNode(ChildNode) Then           '## Yes. Is it a root node?
                oTree.Visible = True
                Exit Function                       '## Yes. Can't move a root node.
            End If
            '## move the children before moving the designated node
            lCount = .Children
            For lLoop = 1 To lCount
                lNDX = .Child.Index
                Set oTree.Nodes(lNDX).Parent = .Parent
            Next
        End If
        '## Force the ParentNode to be expanded before the move
        ParentNode.Expanded = True                  '@@ v01.00.03
        '## Give the child a new parent
        Set .Parent = ParentNode
        If bSelect Then
            .EnsureVisible
            .Selected = bSelect
        End If
    End With
    oTree.Visible = True
    pNodeSwap = True

End Function

Public Function NodeMove(ByVal TargetNode As MSComctlLib.Node, _
                         ByRef SourceNode As MSComctlLib.Node, _
                Optional ByVal bSelect As Boolean = True, _
                Optional ByVal eRelationship As TreeRelationshipConstants = tvwChild) As Boolean

    Dim lLoop    As Long
    Dim lCount   As Long
    Dim oFVNode  As MSComctlLib.Node
    Dim oTmpNode As MSComctlLib.Node
    Dim oNewNode As MSComctlLib.Node

    If (TargetNode Is Nothing) Then Exit Function
    If (SourceNode Is Nothing) Then Exit Function
    If TargetNode.Index = SourceNode.Index Then Exit Function

    If IsParentNode(TargetNode, SourceNode) Then
        NodeMove = pNodeSwap(TargetNode, SourceNode, bSelect)
        Exit Function
    End If

    '
    '## Remember the first visible node
    '
    Set oFVNode = FirstVisibleNode
    '
    '## Hide control whilst moving nodes around to stop flickering
    '
    oTree.Visible = False
    '
    '## 1. Create a temp node for the kids
    '
    If IsRootNode(SourceNode) Then
        With SourceNode
            Set oTmpNode = NodeAdd(, , _
                                   .Key + "[Tmp]", _
                                   .Text, _
                                   .Image, _
                                   .SelectedImage, _
                                   .Tag, _
                                   .Bold, _
                                   .Checked, , , , _
                                   .Sorted, _
                                   .ForeColor, _
                                   .BackColor, _
                                   .ExpandedImage)
        End With
    Else
        With SourceNode
            Set oTmpNode = NodeAdd(.Parent.Key, _
                                   tvwChild, _
                                   .Key + "[Tmp]", _
                                   .Text, _
                                   .Image, _
                                   .SelectedImage, _
                                   .Tag, _
                                   .Bold, _
                                   .Checked, , , , _
                                   .Sorted, _
                                   .ForeColor, _
                                   .BackColor, _
                                   .ExpandedImage)
        End With
    End If
    '
    '## 2. Move the kids to the foster parent
    '
    lCount = SourceNode.Children
    For lLoop = 1 To lCount
        Set oTree.Nodes(SourceNode.Child.Index).Parent = oTmpNode
    Next
    '
    '## 3. Delete the old parent
    '
    oTree.Nodes.Remove SourceNode.Index
    '
    '## 4. Add a new parent
    '
    With oTmpNode
        Set oNewNode = NodeAdd(TargetNode.Key, _
                               eRelationship, _
                               Replace(.Key, "[Tmp]", ""), _
                               .Text, _
                               .Image, _
                               .SelectedImage, _
                               .Tag, _
                               .Bold, _
                               .Checked, , , _
                               True, _
                               .Sorted, _
                               .ForeColor, _
                               .BackColor, _
                               .ExpandedImage)
    End With
    '
    '## 5. Give the kids to the new parent
    '
    lCount = oTmpNode.Children
    For lLoop = 1 To lCount
        Set oTree.Nodes(oTmpNode.Child.Index).Parent = oNewNode
    Next
    '
    '## 6. Delete the foster parent
    '
    oTree.Nodes.Remove oTmpNode.Index
    '
    '## Make sure that the control only scrolls if it's necessary.
    '
    On Error Resume Next    '## If oFVNode = oNewNode then an error will be
                            '   raised as the reference node was deleted but
                            '   VB doesn't destory oFVNode. So the test If
                            '   Not (oFVNode is Nothing) then ... will
                            '   generate an error!
    oFVNode.EnsureVisible
    With oNewNode
        If .Children Then
            .Child.LastSibling.EnsureVisible
        End If
        .EnsureVisible
    End With
    '
    '## Show changes made
    '
    oTree.Visible = True
    NodeMove = True

End Function

Public Function NodeRename(ByVal Node As MSComctlLib.Node, _
                           ByVal NewNodeText As String, _
                  Optional ByVal bSelect As Boolean = True) As Boolean
Attribute NodeRename.VB_Description = "Changes the text of a node and optionally selects & ensures visibility."

    With Node
        If Len(NewNodeText) Then            '## Check if we have new text
            .Text = NewNodeText
            If Not IsRootNode(Node) Then
                .EnsureVisible
                .Selected = bSelect
            End If
            NodeRename = True
        End If
    End With

End Function

Public Sub Redraw(ByVal Enabled As Boolean)
Attribute Redraw.VB_Description = "Enable/disable screen updates."

    Dim hWnd  As Long

    If Not (oTree Is Nothing) Then
        hWnd = oTree.hWnd
        If Enabled Then
            '## Lock the window update to avoid flickering
            SendMessageLong hWnd, WM_SETREDRAW, False, &H0
        Else
            '## Unlock the window
            SendMessageLong hWnd, WM_SETREDRAW, True, &H0
        End If
    End If

End Sub

Public Function Save(ByVal FileName As String, _
            Optional ByVal FileMode As eFileMode) As Boolean
Attribute Save.VB_Description = "Save Treeview to Binary or XML file."
    Select Case FileMode
        Case [Binary]: Save = pSaveBIN(FileName)
        Case [xml]:    Save = pSaveXML(FileName)
    End Select
End Function

Public Sub ScrollView(ByVal Dir As eCodeScrollView)
Attribute ScrollView.VB_Description = "Scrolls the treview using code in all four directions."
    '
    '## Scrolls the treview using code
    '
    Dim lHwnd As Long
    Dim lPos  As Long
    Dim lBar1 As Long
    Dim lBar2 As Long
    Dim lDir  As Long

    lHwnd = oTree.hWnd
    Select Case Dir
        Case [Home]:       SendMessageAny lHwnd, WM_VSCROLL, SB_TOP, vbNull
        Case [Page Up]:    SendMessageAny lHwnd, WM_VSCROLL, SB_PAGEUP, vbNull
        Case [Up]:         SendMessageAny lHwnd, WM_VSCROLL, SB_LINEUP, vbNull
        Case [Down]:       SendMessageAny lHwnd, WM_VSCROLL, SB_LINEDOWN, vbNull
        Case [Page Down]:  SendMessageAny lHwnd, WM_VSCROLL, SB_PAGEDOWN, vbNull
        Case [End]:        SendMessageAny lHwnd, WM_VSCROLL, SB_BOTTOM, vbNull
        Case [Left]:       SendMessageAny lHwnd, WM_HSCROLL, SB_LEFT, vbNull
        Case [Page Left]:  SendMessageAny lHwnd, WM_HSCROLL, SB_PAGELEFT, vbNull
        Case [Line Left]:  SendMessageAny lHwnd, WM_HSCROLL, SB_LINELEFT, vbNull
        Case [Line Right]: SendMessageAny lHwnd, WM_HSCROLL, SB_LINERIGHT, vbNull
        Case [Page Right]: SendMessageAny lHwnd, WM_HSCROLL, SB_PAGERIGHT, vbNull
        Case [Right] ':    SendMessageAny lHwnd, WM_HSCROLL, SB_RIGHT, vbNull
            '
            '## For some reason, the treeview doesn't respond to the above commented
            '   out message. Therefore a work-around is required.
            '
            '## To stop flickering, the control is hidden temporarily.
            oTree.Visible = False
            ' ## Loop until we've scrolled to the far right side
            Do
                lPos = GetScrollPos(lHwnd, 0&)
                SendMessageAny lHwnd, WM_HSCROLL, SB_PAGERIGHT, vbNull
            Loop Until (lPos = GetScrollPos(lHwnd, 0&))
            '## Now show the control
            oTree.Visible = True
    End Select

End Sub

Public Sub SetFocusNode(ByVal KeyIndex As Variant)
Attribute SetFocusNode.VB_Description = "Shows the focus rectangle without actually selecting the node."
    With oTree.Nodes(KeyIndex)
        .Selected = True
        .Selected = False
    End With
End Sub

Public Sub SetOverlayImage(ByVal Imaglelisthandle As OLE_HANDLE, _
                           ByVal ImageListIndex As Long, _
                           ByVal ShareOverlay As eItemOverlayIndices)
Attribute SetOverlayImage.VB_Description = "Initialises Overlay Icon Images. Must be used before setting the Node's OverlayIcon state image."
    If Imaglelisthandle Then
        ImageList_SetOverlayImage Imaglelisthandle, ImageListIndex, ShareOverlay
    End If
End Sub

Public Sub UnHookCtrl()
Attribute UnHookCtrl.VB_Description = "Releases the hooked TreeView control."
    '
    '## Release the hooked control
    '
    Set oTree = Nothing
End Sub

'===========================================================================
' MultiNode Selection: Public properties, subroutines and functions
'
Private Sub iMultiSelect_ClearSelection(Optional ByRef Node As MSComctlLib.INode)

    Dim oNode  As MSComctlLib.Node
    Dim oChild As MSComctlLib.Node
    Dim lPtr   As Long

    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Sub
    End If

    If Node Is Nothing Then
        '## Select the first node as none were passed
        Set oNode = pGetFirstNode()
    Else
        Set oNode = Node
    End If
    If oNode Is Nothing Then
        '## Treeview has no nodes
        Exit Sub
    End If
    Do
        lPtr = ObjPtr(oNode)    '## Tag collection node ID
        If moSelNodes.Exist(lPtr) Then
            '## Node exists so remove node
            pSetNodeSelect oNode, False
        End If
        If moSelNodes.Count Then
            If oNode.Children Then
                '## Node has child nodes, repeat for all child nodes
                Set oNode = oNode.Child
                Do
                    iMultiSelect_ClearSelection oNode
                    If Not (oNode = oNode.LastSibling) Then
                        Set oNode = oNode.Next
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If Not (Node Is Nothing) Then
            Exit Do
        Else
            If Not (oNode = oNode.LastSibling) Then
                Set oNode = oNode.Next
            Else
                Exit Do
            End If
        End If
    Loop
    RaiseEvent SelChange

End Sub

Private Property Get iMultiSelect_FocusNode() As MSComctlLib.INode
    Set iMultiSelect_FocusNode = moFocusNode
End Property

Private Property Let iMultiSelect_NoClearOnSpaceClick(ByVal RHS As Boolean)

    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Property
    End If

    mbSelNoClear = RHS

End Property

Private Property Get iMultiSelect_NoClearOnSpaceClick() As Boolean
    iMultiSelect_NoClearOnSpaceClick = mbSelNoClear
End Property

Private Property Let iMultiSelect_NoDefaultSel(ByVal RHS As Boolean)

    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Property
    End If

    mbSelNoDefault = RHS

End Property

Private Property Get iMultiSelect_NoDefaultSel() As Boolean
    iMultiSelect_NoDefaultSel = mbSelNoDefault
End Property

Private Property Let iMultiSelect_SelBackColor(ByVal RHS As stdole.OLE_COLOR)

    Dim vItem As Variant
    
    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Property
    End If

    Select Case RHS
        Case mlSelBackColor
            '## Already set
        Case Else
            mlSelBackColor = RHS
            On Error Resume Next
            '## Update all selected items
            For Each vItem In moSelNodes
                oTree.Nodes(vItem([NodeKey])).BackColor = mlSelBackColor
                If Err.Number Then
                    Exit Property
                End If
            Next
    End Select

End Property

Private Property Get iMultiSelect_SelBackColor() As stdole.OLE_COLOR
    iMultiSelect_SelBackColor = mlSelBackColor
End Property

Private Property Let iMultiSelect_SelBold(ByVal RHS As Boolean)

    Dim vItem As Variant
    
    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Property
    End If

    Select Case RHS
        Case mbSelBold
            '## Already set
        Case Else
            mbSelBold = RHS
            On Error Resume Next
            '## Update all selected items
            For Each vItem In moSelNodes
                oTree.Nodes(vItem([NodeKey])).Bold = mbSelBold
                If Err.Number Then
                    Exit Property
                End If
            Next
    End Select

End Property

Private Property Get iMultiSelect_SelBold() As Boolean
    iMultiSelect_SelBold = mbSelBold
End Property

Private Sub iMultiSelect_SelectAllNodes(Optional ByRef Node As MSComctlLib.INode)

    Dim oNode  As MSComctlLib.Node
    Dim oChild As MSComctlLib.Node
    Dim lPtr   As Long

    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Sub
    End If

    If Node Is Nothing Then
        '## Select the first node as none were passed
        Set oNode = pGetFirstNode()
    Else
        Set oNode = Node
    End If
    If oNode Is Nothing Then
        '## Treeview has no nodes
        Exit Sub
    End If

    Do
        lPtr = ObjPtr(oNode)    '## Tag collection node ID
        If Not moSelNodes.Exist(lPtr) Then
            '## Node not already selected
            pSetNodeSelect oNode, True
        End If
        If moSelNodes.Count Then
            If oNode.Children Then
                '## Node has child nodes, repeat for all child nodes
                Set oNode = oNode.Child
                Do
                    iMultiSelect_SelectAllNodes oNode
                    If Not (oNode = oNode.LastSibling) Then
                        Set oNode = oNode.Next
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If Not (Node Is Nothing) Then
            Exit Do
        Else
            If Not (oNode = oNode.LastSibling) Then
                Set oNode = oNode.Next
            Else
                Exit Do
            End If
        End If
    Loop
    RaiseEvent SelChange

End Sub

Private Property Get iMultiSelect_SelectedNode(ByVal Index As Long) As Collection

    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Property
    End If

    If moSelNodes.Count Then
        '## Return the requested Node from the collection
        Set iMultiSelect_SelectedNode = oTree.Nodes(moSelNodes(Index, [NodeKey]))
    End If

End Property

Private Property Let iMultiSelect_SelForeColor(ByVal RHS As stdole.OLE_COLOR)

    Dim vItem As Variant
    
    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Property
    End If

    Select Case RHS
        Case mlSelForeColor
            '## Already set
        Case Else
            mlSelForeColor = RHS
            On Error Resume Next
            '## Update all selected items
            For Each vItem In moSelNodes
                oTree.Nodes(vItem([NodeKey])).ForeColor = mlSelForeColor
                If Err.Number Then
                    Exit Property
                End If
            Next
    End Select
    
End Property

Private Property Get iMultiSelect_SelForeColor() As stdole.OLE_COLOR
    iMultiSelect_SelForeColor = mlSelForeColor
End Property

Private Sub iMultiSelect_SelectNode(ByRef Node As MSComctlLib.INode, _
                                    ByVal Selected As Boolean)

    Dim lPtr As Long

    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Sub
    End If

    If Node Is Nothing Then
        Exit Sub
    End If

    pSetNodeSelect Node, Selected
    RaiseEvent SelChange

End Sub

Private Sub iMultiSelect_ToggleNode(Optional ByRef Node As MSComctlLib.INode, _
                                    Optional ByVal Children As Boolean)

    Dim oNode     As MSComctlLib.Node
    Dim oChild    As MSComctlLib.Node
    Dim bSelected As Boolean
    Dim lPtr      As Long

    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Sub
    End If

    If (Node Is Nothing) Then
        '## Select the focus node as none were passed
        Set oNode = moFocusNode
    Else
        Set oNode = Node
    End If
    If oNode Is Nothing Then
        '## Treeview has no nodes
        Exit Sub
    End If

    lPtr = ObjPtr(oNode)    '## Tag collection node ID
    On Error Resume Next
    Select Case oNode.BackColor
        Case mlSelBackColor
            pSetNodeSelect oNode, False
        Case moSelNodes(lPtr, [BackColor])
            pSetNodeSelect oNode, True
            bSelected = True
    End Select
    If moSelNodes.Count Then
        If CBool(oNode.Children) And Children Then
            Set oNode = oNode.Child
            Do While Not (oNode Is Nothing)
                pSelectNode oNode, bSelected, True
                Set oNode = oNode.Next
            Loop
        End If
    End If
    RaiseEvent SelChange

End Sub

Private Sub iMultiSelect_ToggleSelection(Optional ByRef Node As MSComctlLib.INode)

    Dim oNode  As MSComctlLib.Node
    Dim oChild As MSComctlLib.Node
    Dim lPtr   As Long

    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Sub
    End If

    If (Node Is Nothing) Then
        '## Select the first node as none were passed
        Set oNode = pGetFirstNode()
    Else
        Set oNode = Node
    End If
    If oNode Is Nothing Then
        '## Treeview has no nodes
        Exit Sub
    End If

    '## hide the Treeview to speed up the operation
    oTree.Visible = False
    Do
        lPtr = ObjPtr(oNode)    '## Tag collection node ID
        On Error Resume Next
        Select Case oNode.BackColor
            Case mlSelBackColor:                pSetNodeSelect oNode, False
            Case moSelNodes(lPtr, [BackColor]): pSetNodeSelect oNode, True
        End Select
        If moSelNodes.Count Then
            If oNode.Children Then
                Set oNode = oNode.Child
                Do
                    iMultiSelect_ToggleSelection oNode
                    If Not (oNode = oNode.LastSibling) Then
                        Set oNode = oNode.Next
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If Not (Node Is Nothing) Then
            Exit Do
        Else
            If Not (oNode = oNode.LastSibling) Then
                Set oNode = oNode.Next
            Else
                Exit Do
            End If
        End If
    Loop
    oTree.Visible = True
    RaiseEvent SelChange

End Sub

Private Function iMultiSelect_NewEnum() As stdole.IUnknown

    If Not meOpMode = [Multi Select] Then
        '## MultiSelect wasn't enabled. Warn the calling routine and exit
        Err.Raise clTREEMODE_ERR1, csSOURCE_ERR, csTREEMODE_ERR1
        Exit Function
    End If

    '## Pass the enumeration pointers to the Implementation class. Implementation class
    '   will then expose the enumeration to the calling routine. Who said that the
    '   Interface class had no functionality...
    Set iMultiSelect_NewEnum = mSelectedNodes.[_NewEnum]

End Function

'===========================================================================
' MultiNode Selection: Private subroutines and functions
'
Private Function pGetFirstNode() As MSComctlLib.Node
    With oTree.Nodes
        If .Count Then
            Set pGetFirstNode = .Item(1).root.FirstSibling
        End If
    End With
End Function

Private Sub pSelectNode(ByRef Node As MSComctlLib.Node, _
                        ByVal Selected As Boolean, _
               Optional ByVal ChildSelect As Boolean = False)

    Dim oChild As MSComctlLib.Node

    On Error Resume Next
    '## Set select Node state
    pSetNodeSelect Node, Selected
    If ChildSelect Then
        '## Set all child nodes select state as well
        Set oChild = Node.Child
        Do While Not (oChild Is Nothing)
            pSelectNode oChild, Selected, True
            Set oChild = oChild.Next
        Loop
    End If

End Sub

Private Sub pSetNodeSelect(ByRef Node As MSComctlLib.Node, _
                           ByVal State As Boolean)

    Dim lPtr As Long

    lPtr = ObjPtr(Node) '## Tag collection node ID
    With Node
        Select Case State
            Case False
                On Error Resume Next
                '## Get original settings
                .BackColor = moSelNodes(lPtr, [BackColor])
                .ForeColor = moSelNodes(lPtr, [ForeColor])
                .Bold = moSelNodes(lPtr, [Bold])
                If Not Err.Number Then
                    '## No longer required - therefore remove
                    moSelNodes.Remove lPtr
                    mSelectedNodes.Remove CStr(lPtr)
                End If
            Case True
                '## Save original settings before selecting
                moSelNodes.Add lPtr, .Key, .ForeColor, .BackColor, .Bold
                mSelectedNodes.Add Node, CStr(lPtr)
                .BackColor = mlSelBackColor
                .ForeColor = mlSelForeColor
                .Bold = mbSelBold
        End Select
    End With

End Sub

'===========================================================================
' Drag'n'Drop: Private subroutines and functions
'
Private Sub pDoDrag()

    Dim pt         As POINTAPI
    Dim rcClient   As RECT
    Static lOldNdx As Long

    With oTree
        If mbStartDrag = True Then
            If mbInDrag = True Then
                '
                '## If the cursor is still over same item as it was on the previous call,
                '   start the auto expand timer, disable the timer otherwise.
                '
                If Not (.DropHighlight Is Nothing) Then
                    If lOldNdx <> .DropHighlight.Index Then
                        If (.DropHighlight.Children > 0) And (.DropHighlight.Expanded = False) Then
                            moDragExpand.Enabled = True
                        Else
                            moDragExpand.Enabled = False
                        End If
                    End If
                    lOldNdx = .DropHighlight.Index
                End If
                '
                '## If the window is scrollable, and the cursor is within that auto scroll
                '   distance, start the auto scroll timer, disable the timer otherwise.
                '
                GetCursorPos pt
                ScreenToClient .hWnd, pt
                GetClientRect .hWnd, rcClient
                If (pIsWindowScrollable(.hWnd) And pPtInRectRegion(rcClient, mlAutoScroll, pt)) Then
                    moDragScroll.Enabled = True
                Else
                    moDragScroll.Enabled = False
                End If
                If Not (.DropHighlight Is Nothing) Then
                    '## We're over a node
                    RaiseEvent Dragging(moDragNode, .DropHighlight)
                End If
            End If
        End If
    End With

End Sub

'===========================================================================
' General: Private BIN/XML Save/Load subroutines and functions
'
Private Sub pInitXML(ByRef oRs As ADODB.Recordset)
    Set oRs = New ADODB.Recordset
    With oRs.Fields
        .Append "PkID", adInteger, , adFldRowID '## Auto Increment field
        .Append "Text", adVarChar, 255
        .Append "Key", adVarChar, 255
        .Append "Tag", adVariant
        .Append "BackColor", adInteger
        .Append "Bold", adBoolean
        .Append "Checked", adBoolean
        .Append "Expanded", adBoolean
        .Append "ExpandedImage", adInteger
        .Append "ForeColor", adInteger
        .Append "Image", adInteger
        .Append "SelectedImage", adInteger
        .Append "Sorted", adBoolean
        .Append "Selected", adBoolean
        .Append "LinkID", adVarChar, 255
        .Append "FirstVisible", adBoolean
    End With
    oRs.Open
End Sub

Private Function pLoadBIN(ByVal FileName As String) As Boolean

    Dim oNode    As MSComctlLib.Node
    Dim tNodes() As NodeInfo
    Dim fHndl    As Long
    Dim lRecs    As Long
    Dim lLoop    As Long
    Dim sFirst   As String

    On Error GoTo ErrorHandler

    fHndl = FreeFile(0)
    Open FileName For Binary As fHndl
    Get #fHndl, , lRecs
    ReDim tNodes(1 To lRecs)
    Get #fHndl, , tNodes()
    Close #fHndl

    If lRecs Then
        oTree.Nodes.Clear
        For lLoop = 1 To lRecs
            If tNodes(lLoop).LinkId Then
                Set oNode = NodeAdd(tNodes(lLoop).LinkId, _
                                    tvwChild, _
                                    CStr(tNodes(lLoop).Key), _
                                    CStr(tNodes(lLoop).Text), _
                                    CInt(tNodes(lLoop).Image), _
                                    CInt(tNodes(lLoop).SelImage), _
                                    tNodes(lLoop).Tag, _
                                    tNodes(lLoop).Bold, _
                                    tNodes(lLoop).Checked, _
                                    tNodes(lLoop).Sorted, _
                                    tNodes(lLoop).Expanded, _
                                    tNodes(lLoop).Selected, , _
                                    tNodes(lLoop).ForeColor, _
                                    tNodes(lLoop).BackColor, _
                                    CInt(tNodes(lLoop).ExpImage))
            Else
                Set oNode = NodeAdd(, , _
                                    CStr(tNodes(lLoop).Key), _
                                    CStr(tNodes(lLoop).Text), _
                                    CInt(tNodes(lLoop).Image), _
                                    CInt(tNodes(lLoop).SelImage), _
                                    tNodes(lLoop).Tag, _
                                    tNodes(lLoop).Bold, _
                                    tNodes(lLoop).Checked, _
                                    tNodes(lLoop).Sorted, _
                                    tNodes(lLoop).Expanded, _
                                    tNodes(lLoop).Selected, , _
                                    tNodes(lLoop).ForeColor, _
                                    tNodes(lLoop).BackColor, _
                                    CInt(tNodes(lLoop).ExpImage))
            End If
            If tNodes(lLoop).FirstVisible Then
                sFirst = oNode.Key
            End If
        Next
        If Len(sFirst) Then
            oTree.Nodes(sFirst).EnsureVisible
            Set FirstVisibleNode = oTree.Nodes(sFirst)
        End If
        Set oNode = Nothing
    End If
    pLoadBIN = True
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function pLoadXML(ByVal FileName As String) As Boolean

    Dim oNode   As MSComctlLib.Node
    Dim oRs     As ADODB.Recordset
    Dim fHndl   As Long
    Dim sData   As String
    Dim oStream As Stream
    Dim sFirst  As String

    On Error GoTo ErrorHandler

    Set oStream = New Stream
    Set oRs = New ADODB.Recordset
    pInitXML oRs

    fHndl = FreeFile(0)
    Open FileName For Binary As fHndl
    sData = Space$(LOF(fHndl))
    Get #fHndl, , sData
    Close #fHndl

    With oStream
        .Open
        .WriteText sData
        .Position = 0
    End With
    With oRs
        .Close
        .Open oStream
        .MoveFirst
        If .RecordCount Then
            oTree.Nodes.Clear
            Do
                If Len(!LinkId) Then
                    Set oNode = NodeAdd(CStr(!LinkId), _
                                        tvwChild, _
                                        CStr(!Key), _
                                        CStr(!Text), _
                                        CInt(!Image), _
                                        CInt(!SelectedImage), _
                                        !Tag, _
                                        !Bold, _
                                        !Checked, _
                                        !Sorted, _
                                        !Expanded, _
                                        !Selected, , _
                                        !ForeColor, _
                                        !BackColor, _
                                        CInt(!ExpandedImage))
                Else
                    Set oNode = NodeAdd(, , _
                                        CStr(!Key), _
                                        CStr(!Text), _
                                        CInt(!Image), _
                                        CInt(!SelectedImage), _
                                        !Tag, _
                                        !Bold, _
                                        !Checked, _
                                        !Sorted, _
                                        !Expanded, _
                                        !Selected, , _
                                        !ForeColor, _
                                        !BackColor, _
                                        CInt(!ExpandedImage))
                End If
                If !FirstVisible Then
                    sFirst = oNode.Key
                End If
                .MoveNext
            Loop Until .EOF
            If Len(sFirst) Then
                oTree.Nodes(sFirst).EnsureVisible
                Set FirstVisibleNode = oTree.Nodes(sFirst)
            End If
        End If
        .Close
    End With
    oStream.Close
    Set oStream = Nothing
    Set oRs = Nothing
    pLoadXML = True
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function pSaveBIN(ByVal FileName As String) As Boolean

    Dim oNode    As MSComctlLib.Node
    Dim tNodes() As NodeInfo
    Dim fHndl    As Long
    Dim lNDX     As Long

    On Error GoTo ErrorHandler

    ReDim tNodes(1 To oTree.Nodes.Count) As NodeInfo
    For Each oNode In oTree.Nodes
        lNDX = lNDX + 1
        With oNode
            tNodes(lNDX).Index = .Index
            tNodes(lNDX).Text = .Text
            tNodes(lNDX).Key = .Key
            tNodes(lNDX).Tag = .Tag
            tNodes(lNDX).BackColor = .BackColor
            tNodes(lNDX).Bold = .Bold
            tNodes(lNDX).Checked = .Checked
            tNodes(lNDX).Expanded = .Expanded
            tNodes(lNDX).ExpImage = .ExpandedImage
            tNodes(lNDX).ForeColor = .ForeColor
            tNodes(lNDX).Image = .Image
            tNodes(lNDX).SelImage = .SelectedImage
            tNodes(lNDX).Sorted = .Sorted
            tNodes(lNDX).Selected = (.Key = oTree.SelectedItem.Key)
            If IsParentNode(oNode, .Parent) Then
                tNodes(lNDX).LinkId = .Parent.Index
            Else
                tNodes(lNDX).LinkId = 0
            End If
            tNodes(lNDX).FirstVisible = (.Key = FirstVisibleNode.Key)
        End With
    Next

    fHndl = FreeFile(0)
    Open FileName For Binary As fHndl
    Put #fHndl, , CLng(oTree.Nodes.Count)
    Put #fHndl, , tNodes()
    Close #fHndl

    Set oNode = Nothing
    pSaveBIN = True
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function pSaveXML(ByVal FileName As String) As Boolean

    Dim oNode   As MSComctlLib.Node
    Dim oRs     As ADODB.Recordset
    Dim fHndl   As Long
    Dim oStream As Stream
    
    On Error GoTo ErrorHandler

    Set oStream = New Stream
    Set oRs = New ADODB.Recordset
    pInitXML oRs

    With oRs
        For Each oNode In oTree.Nodes
            .AddNew
            !Text = oNode.Text
            !Key = oNode.Key
            !Tag = oNode.Tag
            !BackColor = oNode.BackColor
            !Bold = oNode.Bold
            !Checked = oNode.Checked
            !Expanded = oNode.Expanded
            !ExpandedImage = oNode.ExpandedImage
            !ForeColor = oNode.ForeColor
            !Image = oNode.Image
            !SelectedImage = oNode.SelectedImage
            !Sorted = oNode.Sorted
            !Selected = (oNode.Key = oTree.SelectedItem.Key)
            If IsParentNode(oNode, oNode.Parent) Then
                !LinkId = oNode.Parent.Key
            End If
            !FirstVisible = (oNode.Key = FirstVisibleNode.Key)
            .Update
        Next

        fHndl = FreeFile(0)
        Open FileName For Binary As fHndl
        .Save oStream, adPersistXML
        Put #fHndl, , oStream.ReadText
        Close #fHndl
        .Close
    End With
    oStream.Close
    Set oStream = Nothing
    Set oRs = Nothing
    pSaveXML = True
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'===========================================================================
' General: Private subroutines and functions
'
Private Function pGetNodeHandle(Optional ByVal Node As MSComctlLib.Node) As Long
    If (Node Is Nothing) Then
        pGetNodeHandle = SendMessageAny(oTree.hWnd, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&)
    Else
        pGetNodeHandle = pGetItemFromNode(Node)
    End If
End Function

Private Function pGetItemFromNode(ByVal Node As MSComctlLib.Node) As Long
    '
    '## If successful, returns the treeview item handle represented by
    '   the specified Node, returns 0 otherwise.
    '
    Dim anSiblingPos() As Integer  '## Contains the sibling position of the node and all it's parents
    Dim iLevel         As Integer  '## Hierarchical level of the node
    Dim hitem          As Long
    Dim iLoop          As Integer
    Dim iPos           As Integer

    '
    '## Continually work backwards from the current node to the current node's
    '   first sibling, caching the current node's sibling position in the one-based
    '   array. Then get the first sibling's parent node and start over. Keep going
    '   until the postion of the specified node's top level parent item is obtained...
    '
    Do While (Node Is Nothing) = False
        iLevel = iLevel + 1
        ReDim Preserve anSiblingPos(iLevel)
        anSiblingPos(iLevel) = pGetNodeSiblingPos(Node)
        Set Node = Node.Parent
    Loop

    '## Get the hItem of the first item in the treeview
    hitem = SendMessageAAny(oTree.hWnd, TVM_GETNEXTITEM, ByVal TVGN_ROOT, 0&)
    If hitem Then
        '
        '## Now work backwards through the cached node positions in the array
        '   (from the first treeview node to the specified node), obtaining the respective
        '   item handle for each node at the cached position. When we get to the
        '   specified node's position (the value of the first element in the array), we
        '   got it's hItem...
        '
        For iLoop = iLevel To 1 Step -1
            iPos = anSiblingPos(iLoop)
            Do While iPos > 1
                hitem = SendMessageAAny(oTree.hWnd, TVM_GETNEXTITEM, ByVal TVGN_NEXT, ByVal hitem)
                iPos = iPos - 1
            Loop
            If (iLoop > 1) Then hitem = SendMessageAAny(oTree.hWnd, TVM_GETNEXTITEM, ByVal TVGN_CHILD, ByVal hitem)
        Next
        pGetItemFromNode = hitem
    End If

End Function

Private Function pGetNodeSiblingPos(ByVal Node As MSComctlLib.Node) As Integer
    '
    '## Returns the one-base position of the specified node
    '   with respect to it's sibling order.
    '
    Dim nPos As Integer

    '## Keep counting up from one until the node has no more previous siblings
    Do While (Node Is Nothing) = False
        nPos = nPos + 1
        Set Node = Node.Previous
    Loop
    pGetNodeSiblingPos = nPos

End Function

Private Function pIsWindowScrollable(hWnd As Long) As ScrollDirectionFlags
    '
    '## Returns a set of bit flags indicating whether the specified
    '   window can be scrolled in any given direction.

    Dim si            As SCROLLINFO
    Dim dwScrollFlags As ScrollDirectionFlags

    si.cbSize = Len(si)
    si.fMask = SIF_ALL
    '
    '## Get the horizontal scrollbar's info (GetScrollInfo returns
    '   TRUE after a scrollbar has been added to a window,
    '   even if the respective style bit is not set...)
    '
    If (GetWindowLong(hWnd, GWL_STYLE) And WS_HSCROLL) Then
        If GetScrollInfo(hWnd, SB_HORZ, si) Then
            dwScrollFlags = (sdLeft And (si.nPos > 0))
            dwScrollFlags = dwScrollFlags Or (sdRight And (si.nPos < (((si.nMax - si.nMin) + 1) - si.nPage)))
        End If
    End If
    '
    '## Get the vertical scrollbar's info.
    '
    If (GetWindowLong(hWnd, GWL_STYLE) And WS_VSCROLL) Then
        If GetScrollInfo(hWnd, SB_VERT, si) Then
            dwScrollFlags = dwScrollFlags Or (sdUp And (si.nPos > 0))
            dwScrollFlags = dwScrollFlags Or (sdDown And (si.nPos < (((si.nMax - si.nMin) + 1) - si.nPage)))
        End If
    End If

    pIsWindowScrollable = dwScrollFlags

End Function

Private Function pPtInRectRegion(rc As RECT, cxyRegion As Long, pt As POINTAPI) As RectFlags
    '
    '## Returns a set of bit flags indicating whether the specified point resides in
    '   the specified size region with the perimeter of the specified rect. cxyRegion
    '   defines the rectangular region within rc, and must be a positive value

    Dim dwFlags As RectFlags

    If PtInRect(rc, pt.x, pt.y) Then
        dwFlags = (rfLeft And (pt.x <= (rc.Left + cxyRegion)))
        dwFlags = dwFlags Or (rfRight And (pt.x >= (rc.Right - cxyRegion)))
        dwFlags = dwFlags Or (rfTop And (pt.y <= (rc.Top + cxyRegion)))
        dwFlags = dwFlags Or (rfBottom And (pt.y >= (rc.Bottom - cxyRegion)))
    End If

    pPtInRectRegion = dwFlags

End Function

Private Sub pScale(ByVal InX As Single, ByVal InY As Single, _
                    ByRef OutX As Single, ByRef OutY As Single, _
           Optional ByVal InScaleMode As VBRUN.ScaleModeConstants = vbPixels, _
           Optional ByVal OutScaleMode As VBRUN.ScaleModeConstants = vbTwips)
    '
    ' ## Converts from Pixels to nominated scale mode
    '
    With oTree.Parent
        OutX = .ScaleX(InX, InScaleMode, OutScaleMode)
        OutY = .ScaleY(InY, InScaleMode, OutScaleMode)
    End With

End Sub

'Private Sub pScale(ByVal InX As Single, ByVal InY As Single, _
'                   ByRef OutX As Single, ByRef OutY As Single, _
'                   ByVal ScaleMode As VBRUN.ScaleModeConstants)
'    '
'    ' ## Converts from Pixels to nominated scale mode. Use this
'    '    if host doesn't support ScaleX/ScaleY functions
'    '
'    Dim snTwipX As Single
'    Dim snTwipY As Single
'
'    '## Return calculations
'    snTwipX = InX * Screen.TwipsPerPixelX
'    snTwipY = InY * Screen.TwipsPerPixelY
'
'    Select Case ScaleMode
'        Case vbUser
'        Case vbTwips
'            OutX = snTwipX
'            OutY = snTwipY
'        Case vbPoints
'            OutX = snTwipX / 20
'            OutY = snTwipY / 20
'        Case vbPixels
'            OutX = InX
'            OutY = InY
'        Case vbCharacters
'            OutX = snTwipX / 120
'            OutY = snTwipY / 240
'        Case vbInches
'            OutX = snTwipX / 1440
'            OutY = snTwipY / 1440
'        Case vbMillimeters
'            OutX = snTwipX / 56.7
'            OutY = snTwipY / 56.7
'        Case vbCentimeters
'            OutX = snTwipX / 567
'            OutY = snTwipY / 567
'        Case vbHimetric
'            '##      Twip to Inch to Himetric
'            OutX = snTwipX / 1440 / 2540
'            OutY = snTwipY / 1440 / 2540
'        Case vbContainerPosition
'        Case vbContainerSize
'    End Select
'
'End Sub

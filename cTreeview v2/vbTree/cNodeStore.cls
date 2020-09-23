VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNodeStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *** Den Artikel zu diesem Modul finden Sie unter http://www.aboutvb.de/kom/artikel/komtvwnodestore.htm ***

Private pChildNodes    As Collection

Private pBackColor     As OLE_COLOR
Private pBold          As Boolean
Private pChecked       As Boolean
Private pExpanded      As Boolean
Private pExpandedImage As Variant
Private pForeColor     As OLE_COLOR
Private pImage         As Variant
Private pKey           As String
Private pSelected      As Boolean
Private pSelectedImage As Variant
Private pTag           As Variant
Private pText          As String

Public Property Get BackColor() As OLE_COLOR
    BackColor = pBackColor
End Property

Public Property Let BackColor(New_BackColor As OLE_COLOR)
    pBackColor = New_BackColor
End Property

Public Property Get Bold() As Boolean
    Bold = pBold
End Property

Public Property Let Bold(New_Bold As Boolean)
    pBold = New_Bold
End Property

Public Property Get Checked() As Boolean
    Checked = pChecked
End Property

Public Property Let Checked(New_Checked As Boolean)
    pChecked = New_Checked
End Property

Public Property Get ChildNode(KeyIndex As Variant) As cNodeStore
    If Not (pChildNodes Is Nothing) Then
        Set ChildNode = pChildNodes(KeyIndex)
    End If
End Property

Public Property Get Children() As Long
    If Not (pChildNodes Is Nothing) Then
        Children = pChildNodes.Count
    End If
End Property

Public Property Get Expanded() As Boolean
    Expanded = pExpanded
End Property

Public Property Let Expanded(New_Expanded As Boolean)
    pExpanded = New_Expanded
End Property

Public Property Get ExpandedImage() As Variant
    ExpandedImage = pExpandedImage
End Property

Public Property Let ExpandedImage(New_ExpandedImage As Variant)
    pExpandedImage = New_ExpandedImage
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = pForeColor
End Property

Public Property Let ForeColor(New_ForeColor As OLE_COLOR)
    pForeColor = New_ForeColor
End Property

Public Property Get Image() As Variant
    Image = pImage
End Property

Public Property Let Image(New_Image As Variant)
    pImage = New_Image
End Property

Public Property Get Key() As String
    Key = pKey
End Property

Public Property Let Key(New_Key As String)
    pKey = New_Key
End Property

Public Property Get Selected() As Boolean
    Selected = pSelected
End Property

Public Property Get SelectedImage() As Variant
    SelectedImage = pSelectedImage
End Property

Public Property Let SelectedImage(New_SelectedImage As Variant)
    pSelectedImage = New_SelectedImage
End Property

Public Property Get Tag() As Variant
    If IsObject(pTag) Then
        Set Tag = pTag
    Else
        Tag = pTag
    End If
End Property

Public Property Let Tag(New_Tag As Variant)
    zSetTag New_Tag
End Property

Public Property Set Tag(New_Tag As Variant)
    zSetTag New_Tag
End Property

Private Sub zSetTag(New_Tag As Variant)
    If IsObject(New_Tag) Then
        Set pTag = New_Tag
    Else
        pTag = New_Tag
    End If
End Sub

Public Property Get Text() As String
    Text = pText
End Property

Public Property Let Text(New_Text As String)
    pText = New_Text
End Property

Public Sub Serialize(Node As Variant, Optional TreeView As TreeView, Optional ByVal IncludeChildren As Boolean, Optional ByVal Remove As Boolean, Optional ToPropBag As PropertyBag)
    Dim nNode As Node
    Dim nNodeStore As cNodeStore
    Dim nChildKey As String
    
    If IsObject(Node) Then
        If TypeOf Node Is Node Then
            Set nNode = Node
        End If
    Else
        Set nNode = TreeView.Nodes(Node)
    End If
    If Not (nNode Is Nothing) Then
        With nNode
            pBackColor = .BackColor
            pBold = .Bold
            pChecked = .Checked
            pExpanded = .Expanded
            pExpandedImage = .ExpandedImage
            pForeColor = .ForeColor
            pImage = .Image
            pKey = .Key
            pSelected = .Selected
            pSelectedImage = .SelectedImage
            If IsObject(.Tag) Then
                Set pTag = .Tag
            Else
                pTag = .Tag
            End If
            pText = .Text
            If IncludeChildren Then
                If .Children Then
                    Set pChildNodes = New Collection
                    Set nNode = .Child
                    With pChildNodes
                        Do While Not (nNode Is Nothing)
                            Set nNodeStore = New cNodeStore
                            nNodeStore.Serialize nNode, , True
                            nChildKey = nNode.Key
                            If Len(nChildKey) Then
                                .Add nNodeStore, nChildKey
                            Else
                                .Add nNodeStore
                            End If
                            Set nNode = nNode.Next
                        Loop
                    End With
                End If
            Else
                Set pChildNodes = Nothing
            End If
            If Remove Then
                TreeView.Nodes.Remove .Index
            End If
        End With
        If Not (ToPropBag Is Nothing) Then
            Me.WriteProperties ToPropBag
        End If
    End If
End Sub

Public Function Deserialize(TreeView As TreeView, Optional Relative As Variant, Optional ByVal Relationship As Variant, Optional Key As Variant, Optional ByVal UseOriginalKey As Boolean, Optional ByVal IncludeChildren As Boolean, Optional ByVal KeepExpanded As Boolean, Optional ByVal KeepSelection As Boolean, Optional FromPropBag As PropertyBag) As Node
    Dim nNode As Node
    Dim nNodeStore As cNodeStore
    
    If Not (FromPropBag Is Nothing) Then
        Me.ReadProperties FromPropBag
    End If
    If UseOriginalKey Then
        Set nNode = TreeView.Nodes.Add(Relative, Relationship, pKey, pText)
    Else
        Set nNode = TreeView.Nodes.Add(Relative, Relationship, Key, pText)
    End If
    With nNode
        .BackColor = pBackColor
        .Bold = pBold
        .Checked = pChecked
        .ExpandedImage = pExpandedImage
        .ForeColor = pForeColor
        .Image = pImage
        .SelectedImage = pSelectedImage
        If IsObject(pTag) Then
            Set .Tag = pTag
        Else
            .Tag = pTag
        End If
        .Text = pText
        If IncludeChildren Then
            If Not (pChildNodes Is Nothing) Then
                For Each nNodeStore In pChildNodes
                    nNodeStore.Deserialize TreeView, nNode, tvwChild, , UseOriginalKey, True, KeepExpanded, KeepSelection
                Next
            End If
        End If
        .Expanded = pExpanded And KeepExpanded
        If KeepSelection And pSelected Then
            .Selected = True
            Set TreeView.SelectedItem = nNode
        End If
    End With
    Set Deserialize = nNode
End Function

Public Function CopyNode(Node As Variant, FromTreeView As TreeView, Optional ByVal IncludeChildren As Boolean, Optional ByVal KeepExpanded As Boolean, Optional ToTreeView As TreeView, Optional Relative As Variant, Optional ByVal Relationship As Variant, Optional Key As Variant, Optional ByVal DestNodeExpanded As Boolean) As Node
    Set CopyNode = zCopyMoveNode(Node, FromTreeView, IncludeChildren, KeepExpanded, ToTreeView, Relative, Relationship, Key, , DestNodeExpanded)
End Function

Public Function MoveNode(Node As Variant, FromTreeView As TreeView, Optional ByVal IncludeChildren As Boolean, Optional ByVal KeepExpanded As Boolean, Optional ToTreeView As TreeView, Optional Relative As Variant, Optional ByVal Relationship As Variant, Optional Key As Variant, Optional ByVal UseOriginalKey As Boolean, Optional ByVal DestNodeExpanded As Boolean, Optional ByVal KeepSelection As Boolean) As Node
    Set MoveNode = zCopyMoveNode(Node, FromTreeView, IncludeChildren, KeepExpanded, ToTreeView, Relative, Relationship, Key, UseOriginalKey, DestNodeExpanded, KeepSelection, True)
End Function

Private Function zCopyMoveNode(Node As Variant, FromTreeView As TreeView, Optional ByVal IncludeChildren As Boolean, Optional ByVal KeepExpanded As Boolean, Optional ToTreeView As TreeView, Optional Relative As Variant, Optional ByVal Relationship As Variant, Optional Key As Variant, Optional ByVal UseOriginalKey As Boolean, Optional ByVal DestNodeExpanded As Boolean, Optional ByVal KeepSelection As Boolean, Optional ByVal Remove As Boolean) As Node
    Dim nToTreeView As TreeView
    Dim nNode As Node
    
    Me.Serialize Node, FromTreeView, IncludeChildren, Remove
    If ToTreeView Is Nothing Then
        Set nToTreeView = FromTreeView
    Else
        Set nToTreeView = ToTreeView
    End If
    Set nNode = Me.Deserialize(nToTreeView, Relative, Relationship, Key, UseOriginalKey, IncludeChildren, KeepExpanded, KeepSelection)
    With nNode
        If Not (.Parent Is Nothing) Then
            .Expanded = DestNodeExpanded
        End If
    End With
    Set zCopyMoveNode = nNode
End Function

Public Sub ReadProperties(PropBag As PropertyBag)
    Dim nChildren As Long
    Dim nPropBag As PropertyBag
    Dim nNodeStore As cNodeStore
    Dim nChildKey As String
    Dim l As Long
    
    With PropBag
        pBackColor = .ReadProperty("BackColor")
        pBold = .ReadProperty("Bold")
        pChecked = .ReadProperty("Checked")
        pExpanded = .ReadProperty("Expanded")
        pExpandedImage = .ReadProperty("ExpandedImage")
        pForeColor = .ReadProperty("ForeColor")
        pImage = .ReadProperty("Image")
        pKey = .ReadProperty("Key")
        pSelectedImage = .ReadProperty("SelectedImage")
        pTag = .ReadProperty("Tag")
        pText = .ReadProperty("Text")
        nChildren = .ReadProperty("Children")
        If nChildren Then
            Set pChildNodes = New Collection
            For l = 1 To nChildren
                Set nNodeStore = New cNodeStore
                Set nPropBag = New PropertyBag
                nPropBag.Contents = .ReadProperty("Child" & l)
                nNodeStore.ReadProperties nPropBag
                nChildKey = nNodeStore.Key
                If Len(nChildKey) Then
                    pChildNodes.Add nNodeStore, nChildKey
                Else
                    pChildNodes.Add nNodeStore
                End If
            Next
        End If
    End With
End Sub

Public Sub WriteProperties(PropBag As PropertyBag)
    Dim nNodeStore As cNodeStore
    Dim nPropBag As PropertyBag
    Dim l As Long
    
    With PropBag
        .WriteProperty "BackColor", pBackColor
        .WriteProperty "Bold", pBold
        .WriteProperty "Checked", pChecked
        .WriteProperty "Expanded", pExpanded
        .WriteProperty "ExpandedImage", pExpandedImage
        .WriteProperty "ForeColor", pForeColor
        .WriteProperty "Image", pImage
        .WriteProperty "Key", pKey
        .WriteProperty "SelectedImage", pSelectedImage
        If Not IsObject(pTag) Then
            .WriteProperty "Tag", pTag
        End If
        .WriteProperty "Text", pText
        If pChildNodes Is Nothing Then
            .WriteProperty "Children", 0
        Else
            .WriteProperty "Children", pChildNodes.Count
            For Each nNodeStore In pChildNodes
                Set nPropBag = New PropertyBag
                nNodeStore.WriteProperties nPropBag
                l = l + 1
                .WriteProperty "Child" & l, nPropBag.Contents
            Next
        End If
    End With
End Sub

Public Function Add(Optional Key As Variant, Optional Before As Variant, Optional After As Variant) As cNodeStore
    Dim nNodeStore As cNodeStore
    
    If pChildNodes Is Nothing Then
        Set pChildNodes = New Collection
    End If
    Set nNodeStore = New cNodeStore
    pChildNodes.Add nNodeStore, Key, Before, After
    Set Add = nNodeStore
End Function

Public Sub Clear()
    Set pChildNodes = Nothing
End Sub

Public Sub Remove(KeyIndex As Variant)
    If Not (pChildNodes Is Nothing) Then
        pChildNodes.Remove KeyIndex
    End If
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = pChildNodes.[_NewEnum]
End Function


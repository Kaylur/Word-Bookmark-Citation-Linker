VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBookmarkPick 
   Caption         =   "Bookmark Wizard"
   ClientHeight    =   5820
   ClientLeft      =   240
   ClientTop       =   945
   ClientWidth     =   9930.001
   OleObjectBlob   =   "frmBookmarkPick.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmBookmarkPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SelectedBookmark As String
Public Cancelled As Boolean
Public Action As String
Public CanGoBack As Boolean

Private Sub lblContext_Click()

End Sub

'========================
' INIT
'========================
Private Sub UserForm_Initialize()

    Dim bm As Bookmark

    cmbBookmarks.Clear

    SelectedBookmark = ""
    Cancelled = False
    Action = ""
    CanGoBack = False

    For Each bm In ActiveDocument.Bookmarks
        cmbBookmarks.AddItem bm.Name
    Next bm

    If cmbBookmarks.ListCount > 0 Then
        cmbBookmarks.ListIndex = -1
    End If

    lblCitation.Caption = "Citation: (none)"
    lblContext.Caption = ""

    cmdBack.Enabled = False

    Me.StartUpPosition = 0
    Me.Left = Application.Left + 50
    Me.Top = Application.Top + 50

End Sub

'========================
' LINK (ONE CLICK)
'========================
Private Sub cmbBookmarks_Change()

    If cmbBookmarks.Value = "" Then Exit Sub

    SelectedBookmark = cmbBookmarks.Value
    UpdateBookmarkPreview
    Action = "LINK"

End Sub

'========================
' NEXT
'========================
Private Sub cmdNext_Click()

    Action = "NEXT"

End Sub

'========================
' BACK
'========================
Private Sub cmdBack_Click()

    If Not CanGoBack Then Exit Sub
    Action = "BACK"

End Sub

'========================
' CANCEL SESSION
'========================
Private Sub cmdCancel_Click()

    Cancelled = True
    Action = "CANCEL"

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then
        Cancelled = True
    End If

End Sub

'========================
' SHOW BOOKMARK PREVIEW
'========================
'========================
' SHOW BOOKMARK PREVIEW
'========================
Private Sub UpdateBookmarkPreview()

    Dim bmRange As Range
    Dim previewRange As Range
    Dim searchRange As Range
    Dim previewText As String

    Dim startPos As Long
    Dim endPos As Long

    If cmbBookmarks.Value = "" Then
        lblContext.Caption = ""
        Exit Sub
    End If

    If Not ActiveDocument.Bookmarks.Exists(cmbBookmarks.Value) Then
        lblContext.Caption = ""
        Exit Sub
    End If

    Set bmRange = ActiveDocument.Bookmarks(cmbBookmarks.Value).Range
    
    '====================================
    ' TABLE
    '====================================
    If bmRange.Information(wdWithInTable) Then
    
        Dim cellText As String
        cellText = bmRange.Cells(1).Range.Text
    
        ' remove end-of-cell marker (very important)
        cellText = Replace(cellText, Chr(7), "")
    
        cellText = Replace(cellText, vbCr, " ")
        cellText = Replace(cellText, vbTab, " ")
        cellText = Trim(cellText)
    
        If Len(cellText) > 300 Then
            cellText = Left(cellText, 300) & "..."
        End If
    
        lblContext.Caption = cellText
        Exit Sub
    
    End If
    '====================================
    ' FIND PREVIOUS LINE/PARAGRAPH BREAK
    '====================================
    Set searchRange = ActiveDocument.Range(0, bmRange.Start)

    With searchRange.Find
        .ClearFormatting
        .Text = "^p"
        .Forward = False
        .Wrap = wdFindStop
    End With

    If searchRange.Find.Execute Then
        startPos = searchRange.End
    Else
        startPos = 0
    End If

    '====================================
    ' BUILD PREVIEW
    '====================================
    Dim docEnd As Long
    docEnd = ActiveDocument.Content.End
    
    endPos = startPos + 300
    If endPos > docEnd Then endPos = docEnd
    
    Set previewRange = ActiveDocument.Range(startPos + 1, endPos)

    previewText = previewRange.Text
    previewText = Replace(previewText, Chr(7), "") 'table end-of-cell marker
    previewText = Replace(previewText, vbCr, " ")
    previewText = Replace(previewText, vbTab, " ")
    previewText = Trim(previewText)

    If Len(previewText) > 300 Then
        previewText = Left(previewText, 300) & "..."
    End If

    lblContext.Caption = previewText

End Sub

Public Sub RefreshBackState(ByVal historyCount As Long)
    CanGoBack = (historyCount > 0)
    cmdBack.Enabled = CanGoBack
End Sub

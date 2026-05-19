Attribute VB_Name = "BookmarkWizard"
Option Explicit

Sub CitationNavigator()

    Dim doc As Document
    Dim r As Range
    Dim foundRange As Range
    Dim frm As frmBookmarkPick

    Dim docEnd As Long
    Dim txt As String
    Dim beforeText As String
    Dim afterText As String

    Dim citationHistory As Collection
    Dim resumePos As Long
    Dim moveToPos As Long

    Set doc = ActiveDocument
    Set frm = New frmBookmarkPick
    Set citationHistory = New Collection

    docEnd = doc.Content.End

    frm.Show vbModeless

    Set r = doc.Content
    r.Collapse wdCollapseStart

    Do While r.Start < docEnd

        '========================
        ' FIND CITATION
        '========================
        With r.Find
            .ClearFormatting
            .Text = "\([0-9]@\)"
            .MatchWildcards = True
            .Forward = True
            .Wrap = wdFindStop
        End With

        If Not r.Find.Execute Then Exit Do

        Set foundRange = r.Duplicate
        txt = foundRange.Text

        '========================
        ' VALIDATION
        '========================
        If Not (txt Like "([0-9])" Or _
                txt Like "([0-9][0-9])" Or _
                txt Like "([0-9][0-9][0-9])") Then
            GoTo AdvanceOnly
        End If

        '========================
        ' CONTEXT FILTER
        '========================
        If foundRange.Start > 5 Then
            beforeText = doc.Range(foundRange.Start - 5, foundRange.Start).Text
        Else
            beforeText = ""
        End If

        If foundRange.End + 6 <= doc.Content.End Then
            afterText = doc.Range(foundRange.End, foundRange.End + 6).Text
        Else
            afterText = ""
        End If

        If beforeText Like "*###*" Or _
           afterText Like " ###-*" Or _
           afterText Like "*-###*" Then
            GoTo AdvanceOnly
        End If

        '========================
        ' FORMAT CITATION
        '========================
        With foundRange.Font
            .Name = "Georgia"
            .Size = 8
        End With
        
        '========================
        ' UPDATE UI
        '========================
        
        Dim ctx As Range
        Dim previewText As String
        Dim re As Object
        
        Set ctx = doc.Range(foundRange.Start, foundRange.Start)
        ctx.MoveStart wdWord, -8
        
        previewText = ctx.Text
        previewText = Replace(previewText, vbCr, " ")
        previewText = Replace(previewText, vbTab, " ")
        previewText = Trim(previewText)
        
        '========================
        ' REMOVE PRIOR CITATION NUMBERS ONLY
        '========================
        Set re = CreateObject("VBScript.RegExp")
        
        With re
            .Global = True
            .IgnoreCase = True
            .Pattern = "\([0-9]{1,3}\)"
        End With
        
        previewText = re.Replace(previewText, "")
        previewText = Trim(previewText)
        
        ' Clean doubled spaces after removals
        Do While InStr(previewText, "  ") > 0
            previewText = Replace(previewText, "  ", " ")
        Loop
        
        frm.lblCitation.Caption = "Citation: ..." & previewText & " " & foundRange.Text

        'frm.lblCitation.Caption = "Citation: " & foundRange.Text
        'frm.lblContext.Caption = ""

        frm.RefreshBackState citationHistory.count

        '========================
        ' RESET STATE
        '========================
        frm.Action = ""
        frm.SelectedBookmark = ""
        frm.Cancelled = False

        '========================
        ' WAIT FOR USER INPUT
        '========================
        frm.RefreshBackState citationHistory.count
        
        Do While frm.Action = ""
            DoEvents
        Loop

        If frm.Cancelled Then Exit Do

        '========================
        ' PROCESS ACTION
        '========================
        Select Case frm.Action

            Case "LINK"

                If frm.SelectedBookmark <> "" Then

                    If doc.Bookmarks.Exists(frm.SelectedBookmark) Then

                        If foundRange.Hyperlinks.count > 0 Then
                            foundRange.Hyperlinks(1).Delete
                        End If

                        doc.Hyperlinks.Add _
                            Anchor:=foundRange, _
                            Address:="", _
                            SubAddress:=frm.SelectedBookmark, _
                            TextToDisplay:=foundRange.Text

                        'citationHistory.Add foundRange.Start

                    End If

                End If

                ' LINK = treat as NEXT after processing
                GoTo AdvanceOnly

            Case "NEXT"
                citationHistory.Add foundRange.Start
                
                GoTo AdvanceOnly

            Case "BACK"

                If citationHistory.count > 0 Then
            
                    resumePos = citationHistory(citationHistory.count)
                    citationHistory.Remove citationHistory.count
            
                    r.Start = resumePos
                    r.End = docEnd
                    r.Collapse wdCollapseStart
                End If
            
                frm.RefreshBackState citationHistory.count
                GoTo ContinueLoop

            Case "CANCEL"
                Exit Do

        End Select

AdvanceOnly:
        ' ONLY ONE PLACE ADVANCES THE LOOP (CRITICAL FIX)

        r.Start = foundRange.End
        r.End = docEnd
        r.Collapse wdCollapseStart

ContinueLoop:
        DoEvents

    Loop

CleanExit:

    On Error Resume Next
    Unload frm
    Set frm = Nothing
    Set r = Nothing
    Set foundRange = Nothing
    Set doc = Nothing

    MsgBox "Finished linking citations.", vbInformation

End Sub

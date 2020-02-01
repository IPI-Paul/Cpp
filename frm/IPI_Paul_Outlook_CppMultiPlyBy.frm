VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IPI_Paul_Outlook_CppMultiPlyBy 
   Caption         =   "IPI Paul - C++ Multiply By"
   ClientHeight    =   5280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9255.001
   OleObjectBlob   =   "IPI_Paul_Outlook_CppMultiPlyBy.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "IPI_Paul_Outlook_CppMultiPlyBy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rIdx As Integer, cIdx As Integer
Const wdOrientLandscape As Integer = 1, wdOrientPortrait As Integer = 0, wdActiveEndPageNumber As Integer = 3
Const wdBorderLeft As Integer = -2, wdBorderRight As Integer = -4, wdBorderVertical As Integer = -6, wdLineStyleNone As Integer = 0, wdLineStyleSingle As Integer = 1
Const wdAlignParagraphRight As Integer = 2, wdAlignParagraphCenter As Integer = 1, wdAlignParagraphLeft As Integer = 0

Private Function buildHTML()
    styl = Space(2) & "<style>" & vbCrLf
    styl = styl & Space(4) & "table, th, tr {" & vbCrLf
    styl = styl & Space(6) & "border-collapse: collapse;" & vbCrLf
    styl = styl & Space(4) & "}" & vbCrL
    styl = styl & Space(4) & "table, th, tr, td {" & vbCrLf
    styl = styl & Space(6) & "padding: 0px 5px 0px 5px;" & vbCrLf
    styl = styl & Space(6) & "border: 1 solid black;" & vbCrLf
    styl = styl & Space(6) & "font-size: 1em;" & vbCrLf
    styl = styl & Space(4) & "}" & vbCrLf
    styl = styl & Space(4) & "th {" & vbCrLf
    styl = styl & Space(6) & "color: rgb(" & cboFgRed & ", " & cboFgGreen & ", " & cboFgBlue & ");" & vbCrLf
    styl = styl & Space(6) & "background-color: rgb(" & cboBgRed & ", " & cboBgGreen & ", " & cboBgBlue & ");" & vbCrLf
    styl = styl & Space(4) & "}" & vbCrLf
    styl = styl & Space(4) & "td {" & vbCrLf
    styl = styl & Space(6) & "text-align: right;" & vbCrLf
    styl = styl & Space(4) & "}" & vbCrLf
    styl = styl & Space(2) & "</style>" & vbCrLf
    tbl = Space(2) & "<table id=""multiplyBy"">" & vbCrLf
    tr = ""
    trd = Space(4) & "<tr>" & vbCrLf
    tre = Space(4) & "</tr>" & vbCrLf
    thd = Space(6) & "<th>" & vbCrLf & Space(8)
    the = vbCrLf & Space(6) & "</th>" & vbCrLf
    tdd = Replace(thd, "th", "td")
    tde = Replace(the, "th", "td")
    For i = 0 To lbo_MultiplyBy.ListCount - 1
        th = ""
        td = ""
        For j = 0 To lbo_MultiplyBy.ColumnCount - 1
            If i = 0 Then
                th = th & thd & lbo_MultiplyBy.List(i, j) & the
            Else
                td = td & tdd & restyle(2, lbo_MultiplyBy.List(i, j)) & tde
            End If
        Next j
        If i = 0 Then
            tr = tr & trd & th & tre
        Else
            tr = tr & trd & td & tre
        End If
    Next i
    tr = tr & trd & Space(6) & "<td colspan=""3"" style=""border-left: none; border-right: none;"">  &nbsp; </td>" & tre
    td = ""
    For i = 0 To lboTotals.ColumnCount - 1
        td = td & tdd & lboTotals.List(0, i) & tde
    Next i
    tr = tr & trd & td & tre
    htm = styl & tbl & tr & Space(2) & "</table>"
    buildHTML = htm
End Function

Private Sub buildWord(doc)
    Set tbl = doc.Application.Selection.Tables.Add(doc.Application.Selection.Range(), lbo_MultiplyBy.ListCount + 2, lbo_MultiplyBy.ColumnCount)
    With tbl.Borders
        .InsideLineStyle = wdLineStyleSingle
        .OutsideLineStyle = wdLineStyleSingle
        .InsideColor = rgb(cboBgRed, cboBgGreen, cboBgBlue)
        .OutsideColor = rgb(cboBgRed, cboBgGreen, cboBgBlue)
    End With
    pg = tbl.Rows(1).Range.Information(wdActiveEndPageNumber)
    rws = 1
    For i = 0 To lbo_MultiplyBy.ListCount - 1
        If pg < tbl.Rows(rws).Range.Information(wdActiveEndPageNumber) Then
            pg = tbl.Rows(rws).Range.Information(wdActiveEndPageNumber)
            tbl.Rows.Add (tbl.Rows(rws))
            For j = 0 To lbo_MultiplyBy.ColumnCount - 1
                With tbl.Rows(rws).Cells(j + 1).Range
                    .Text = lbo_MultiplyBy.List(0, j)
                    .Shading.BackgroundPatternColor = rgb(cboBgRed, cboBgGreen, cboBgBlue)
                    .Font.Color = rgb(cboFgRed, cboFgGreen, cboFgBlue)
                    .Font.Bold = True
                End With
            Next j
            rws = rws + 1
        End If
        If i = 0 Then
            For j = 0 To lbo_MultiplyBy.ColumnCount - 1
                With tbl.Rows(rws).Cells(j + 1).Range
                    .Text = lbo_MultiplyBy.List(0, j)
                    .Shading.BackgroundPatternColor = rgb(cboBgRed, cboBgGreen, cboBgBlue)
                    .Font.Color = rgb(cboFgRed, cboFgGreen, cboFgBlue)
                    .Font.Bold = True
                End With
            Next j
        Else
            For j = 0 To lbo_MultiplyBy.ColumnCount - 1
                tbl.Rows(rws).Cells(j + 1).Range.Text = restyle(2, lbo_MultiplyBy.List(i, j))
                tbl.Rows(rws).Cells(j + 1).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
            Next j
        End If
        rws = rws + 1
    Next i
    With tbl.Rows(rws).Cells
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
    End With
    rws = rws + 1
    For j = 0 To lboTotals.ColumnCount - 1
        tbl.Rows(rws).Cells(j + 1).Range.Text = restyle(2, lboTotals.List(0, j))
        tbl.Rows(rws).Cells(j + 1).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
    Next j
    Set tbl = Nothing
End Sub

Private Sub cboBgBlue_Change()
    updColours
End Sub

Private Sub cboBgGreen_Change()
    updColours
End Sub

Private Sub cboBgRed_Change()
    updColours
End Sub

Private Sub cboFgBlue_Change()
    updColours
End Sub

Private Sub cboFgGreen_Change()
    updColours
End Sub

Private Sub cboFgRed_Change()
    updColours
End Sub

Private Sub cboFormats_Change()
    For i = 1 To lbo_MultiplyBy.ListCount - 1
        For j = 0 To lbo_MultiplyBy.ColumnCount - 1
            lbo_MultiplyBy.List(i, j) = restyle(1, lbo_MultiplyBy.List(i, j))
        Next j
    Next i
    updTotals
End Sub

Private Sub cboFunctions_Change()
    If Not cboFunctions.ListIndex = 0 Then
        If cboFunctions.Text = "Send to New Email" Or cboFunctions.Text = "Send to Open Email" Then
            htm = buildHTML
            If cboFunctions.Text = "Send to New Email" Then
                Set oMail = Application.CreateItem(olMailItem)
                oMail.Display
                oMail.Subject = "C++ multiplyBy Outlook Listbox Test"
                oMail.HTMLBody = "<br /><br />" & htm
            Else
                Set oMail = Application.ActiveInspector().CurrentItem
                oMail.Display
                Application.ActiveInspector().WordEditor.Application.Selection = "placeHere"
                oMail.HTMLBody = Replace(oMail.HTMLBody, "placeHere", htm)
            End If
            Set oMail = Nothing
        ElseIf cboFunctions.Text = "Send to New Word Doc" Or cboFunctions.Text = "Send to Open Word Doc" Then
            If cboFunctions.Text = "Send to New Word Doc" Then
                Set wrd = CreateObject("Word.Application")
                wrd.Visible = True
                Set doc = wrd.Documents.Add()
                doc.PageSetup.Orientation = wdOrientLandscape
                With doc.PageSetup
                    .TopMargin = CentimetersToPoints(1.75)
                    .LeftMargin = CentimetersToPoints(1.75)
                    .BottomMargin = CentimetersToPoints(1.75)
                    .RightMargin = CentimetersToPoints(1.75)
                End With
            Else
                Set wrd = GetObject(, "Word.Application")
                If wrd.Documents.Count = 0 Then
                    wrd.Quit
                    Set wrd = GetObject(, "Word.Application")
                End If
                Set doc = wrd.ActiveDocument
            End If
            buildWord doc
            Set doc = Nothing
            Set wrd = Nothing
        End If
        cboFunctions.ListIndex = 0
    End If
End Sub

Private Function CentimetersToPoints(centimeter)
    CentimetersToPoints = centimeter * 28.3464567
End Function

Private Sub lbo_MultiplyBy_Click()
    If lbo_MultiplyBy.ListIndex = 0 Then
        rIdx = 1
    Else
        rIdx = lbo_MultiplyBy.ListIndex
    End If
    For i = 1 To 3
        Me.Controls("txtCol" & i).Text = lbo_MultiplyBy.List(rIdx, i - 1)
    Next
End Sub

Private Sub lbo_MultiplyBy_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 3 Then
        Dim objData As DataObject
        txt = ""
        For i = 0 To lbo_MultiplyBy.ListCount - 1
            For j = 0 To lbo_MultiplyBy.ColumnCount - 1
                If j > 0 Then txt = txt & Chr(9)
                If i > 0 Then
                    txt = txt & restyle(2, lbo_MultiplyBy.List(i, j))
                Else
                    txt = txt & lbo_MultiplyBy.List(i, j)
                End If
            Next j
            txt = txt & vbCrLf
        Next i
        txt = txt & vbCrLf
        For i = 0 To lboTotals.ColumnCount - 1
            If i > 0 Then txt = txt & Chr(9)
            txt = txt & lboTotals.List(0, i)
        Next i
        Set objData = New DataObject
        objData.SetText txt
        objData.PutInClipboard
        Set objData = Nothing
    End If
End Sub

Private Function restyle(typ, val)
    If cboFormats.ListIndex = 0 Then
        restyle = normVal(val)
    ElseIf typ = 1 Then
        If cboFormats.ListIndex = 1 Then
            frmt = Replace(cboFormats.Text, "0.00", "0.####################")
        Else
            frmt = Replace(cboFormats.Text, "0", "0.####################")
        End If
        restyle = Format(normVal(val), frmt)
    ElseIf typ = 2 Then
        If cboFormats.ListIndex = 1 Then
            dec = 2
        Else
            dec = 0
        End If
        restyle = Format(Round(normVal(val), dec), cboFormats.Text)
    End If
End Function

Private Function rgb(red, green, blue)
    If red = "" Then red = 0
    If green = "" Then green = 0
    If blue = "" Then blue = 0
    rgb = ((red) + ((green) * 256) + ((blue) * 65536))
End Function

Private Sub txtCol1_AfterUpdate()
    cIdx = 0
    updListbox txtCol1.Value
End Sub

Private Sub txtCol1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    rIdx = lbo_MultiplyBy.ListCount
    lbo_MultiplyBy.AddItem "0", rIdx
    lbo_MultiplyBy.List(rIdx, 1) = "0"
    lbo_MultiplyBy.List(rIdx, 2) = "0"
    lbo_MultiplyBy.ListIndex = rIdx
    For i = 1 To 3
        Me.Controls("txtCol" & i).Text = 0
    Next
End Sub

Private Sub txtCol2_AfterUpdate()
    cIdx = 1
    updListbox txtCol2.Value
End Sub

Private Sub updColours()
    lblHeader.BackColor = rgb(cboBgRed, cboBgGreen, cboBgBlue)
    lblFore.BackColor = rgb(cboBgRed, cboBgGreen, cboBgBlue)
    lblHeader.ForeColor = rgb(cboFgRed, cboFgGreen, cboFgBlue)
    lblFore.ForeColor = rgb(cboFgRed, cboFgGreen, cboFgBlue)
End Sub

Private Sub updListbox(val)
    lbo_MultiplyBy.List(rIdx, cIdx) = restyle(1, val)
    num = normVal(lbo_MultiplyBy.List(rIdx, 0))
    num1 = normVal(lbo_MultiplyBy.List(rIdx, 1))
    res = cppMultiplyBy(CDbl(num), CDbl(num1))
    lbo_MultiplyBy.List(rIdx, 2) = res
    txtCol3 = restyle(1, res)
    updTotals
End Sub

Private Sub updTotals()
    col1 = 0
    col2 = 0
    col3 = 0
    For i = 1 To lbo_MultiplyBy.ListCount - 1
        col1 = col1 + normVal(lbo_MultiplyBy.List(i, 0))
        col2 = col2 + normVal(lbo_MultiplyBy.List(i, 1))
        col3 = col3 + normVal(lbo_MultiplyBy.List(i, 2))
    Next
    lboTotals.List(0, 0) = restyle(2, col1)
    lboTotals.List(0, 1) = restyle(2, col2)
    lboTotals.List(0, 2) = restyle(2, col3)
End Sub

Private Sub UserForm_Initialize()
    updColours
    For Each itm In Array("General", "#,###,##0.00", "#,###,##0")
        cboFormats.AddItem itm, cboFormats.ListCount
    Next
    For Each itm In Array("", "Send to New Email", "Send to Open Email", "Send to New Word Doc", "Send to Open Word Doc")
        cboFunctions.AddItem itm, cboFunctions.ListCount
    Next
    lbo_MultiplyBy.AddItem "Number1", 0
    lbo_MultiplyBy.List(0, 1) = "Number2"
    lbo_MultiplyBy.List(0, 2) = "Multiplied Result"
    lbo_MultiplyBy.AddItem "0", 1
    lbo_MultiplyBy.List(1, 1) = "0"
    lbo_MultiplyBy.List(1, 2) = "0"
    lbo_MultiplyBy.ListIndex = 1
    rIdx = 1
    For i = 1 To 3
        Me.Controls("txtCol" & i).Text = 0
    Next
    lboTotals.AddItem "0", 0
    lboTotals.List(0, 1) = "0"
    lboTotals.List(0, 2) = "0"
    For i = 0 To 255
        cboBgRed.AddItem i, i
        cboBgGreen.AddItem i, i
        cboBgBlue.AddItem i, i
        cboFgRed.AddItem i, i
        cboFgGreen.AddItem i, i
        cboFgBlue.AddItem i, i
    Next i
End Sub


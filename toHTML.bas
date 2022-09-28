Attribute VB_Name = "toHTML"
Option Explicit
Option Base 1

Sub ConvertHTML()
Attribute ConvertHTML.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.Macro2"
'
' This macro converts a simple Word document into clean HTML to include as a blog <section>.
'
'
Dim aWDoc As Document, p As Paragraph, BlogDoc As String
Dim HTML() As String, quote As String, str As String, fileNm As String, longNm As String, LT As String
Dim x As Integer, i As Integer, listCounter As Integer

Application.ScreenUpdating = False

Set aWDoc = ActiveDocument
listCounter = 1
x = aWDoc.Paragraphs.Count

fileNm = aWDoc.Paragraphs(1).Range.Text
fileNm = StrConv(fileNm, 3)
fileNm = Left(fileNm, Len(fileNm) - 1)
fileNm = Replace(fileNm, " ", "")
fileNm = Replace(fileNm, "!", "")
fileNm = Replace(fileNm, "?", "")
fileNm = Replace(fileNm, ".", "")
longNm = ActiveDocument.Path + "\" + fileNm + ".html"

Call PicsAndLinks
Call Formats

Application.ScreenRefresh

ReDim HTML(x + 3) As String

quote = """"
HTML(1) = "<!--This is new blog section-->"
HTML(2) = "<section id=" + quote + fileNm + quote + " class=" + quote + "blog" + quote + ">"

For i = 1 To x
    If aWDoc.Paragraphs(i).Style = "Normal" Then
        str = aWDoc.Paragraphs(i).Range.Text
        str = Left(str, Len(str) - 1)
        HTML(i + 2) = "<p>" & str & "</p>"
    ElseIf aWDoc.Paragraphs(i).Style = "Heading 1" Then
        str = aWDoc.Paragraphs(i).Range.Text
        str = Left(str, Len(str) - 1)
        HTML(i + 2) = "<h1>" & str & "</h1>"
    ElseIf aWDoc.Paragraphs(i).Style = "Heading 2" Then
        str = aWDoc.Paragraphs(i).Range.Text
        str = Left(str, Len(str) - 1)
        HTML(i + 2) = "<h2>" & str & "</h2>"
    ElseIf aWDoc.Paragraphs(i).Style = "Heading 3" Then
        str = aWDoc.Paragraphs(i).Range.Text
        str = Left(str, Len(str) - 1)
        HTML(i + 2) = "<h3>" & str & "</h3>"
    ElseIf aWDoc.Paragraphs(i).Style = "List Paragraph" Then
        If aWDoc.Lists(listCounter).Range.ListFormat.ListType = 2 Then
            LT = "u"
        ElseIf aWDoc.Lists(listCounter).Range.ListFormat.ListType = 3 Then
            LT = "o"
        End If
        If aWDoc.Paragraphs(i).Style = "List Paragraph" And aWDoc.Paragraphs(i - 1).Style = "List Paragraph" And aWDoc.Paragraphs(i + 1).Style = "List Paragraph" Then
            str = aWDoc.Paragraphs(i).Range.Text
            str = Left(str, Len(str) - 1)
            HTML(i + 2) = "<li>" & str & "</li>"
        ElseIf aWDoc.Paragraphs(i).Style = "List Paragraph" And aWDoc.Paragraphs(i - 1).Style <> "List Paragraph" Then
            str = aWDoc.Paragraphs(i).Range.Text
            str = Left(str, Len(str) - 1)
            HTML(i + 2) = "<" & LT & "l><li>" & str & "</li>"
        ElseIf aWDoc.Paragraphs(i).Style = "List Paragraph" And aWDoc.Paragraphs(i + 1).Style <> "List Paragraph" Then
            str = aWDoc.Paragraphs(i).Range.Text
            str = Left(str, Len(str) - 1)
            HTML(i + 2) = "<li>" & str & "</li></" & LT & "l>"
            listCounter = listCounter + 1
        End If
    End If
Next i
HTML(x + 3) = "</section>"

Open longNm For Output As #1
For i = 1 To x + 3
    Print #1, HTML(i)
Next i
Close #1

aWDoc.Close SaveChanges:=wdDoNotSaveChanges

End Sub

Sub Formats()
'
' This macro is used by ConvertHTML() for various misc. formats(strong, em, and whatnot).
'
'
Dim p As Paragraph, aWDoc As Document, wrd As Variant
Dim track As Boolean
Set aWDoc = ActiveDocument

For Each p In aWDoc.Paragraphs
    track = False
    If p.Style = "Normal" Or p.Style = "List Paragraph" Then
        For Each wrd In p.Range.Words
            wrd.Select
            If Selection.Font.Bold = True And track = False Then
                Selection.InsertBefore " <strong> "
                track = True
            ElseIf Selection.Font.Bold = False And track = True Then
                Selection.InsertBefore " </strong> "
                track = False
            End If
        Next wrd
    End If
Next p

For Each p In aWDoc.Paragraphs
    track = False
    If p.Style = "Normal" Or p.Style = "List Paragraph" Then
        For Each wrd In p.Range.Words
            wrd.Select
            If Selection.Font.Italic = True And track = False Then
                Selection.InsertBefore " <em> "
                track = True
            ElseIf Selection.Font.Italic = False And track = True Then
                Selection.InsertBefore " </em> "
                track = False
            End If
        Next wrd
    End If
Next p

aWDoc.Select
    With Selection.Find
        .Text = "  "
        .Replacement.Text = " "
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "> <"
        .Replacement.Text = "><"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = " </"
        .Replacement.Text = "</"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = " <"
        .Replacement.Text = "<"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = " ."
        .Replacement.Text = "."
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = " !"
        .Replacement.Text = "!"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = " ?"
        .Replacement.Text = "?"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "'"
        .Replacement.Text = "&#39;"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub

Sub PicsAndLinks()
'
' This macro is used by ConvertHTML() to pre-format pics and links into HTML tags (pictures in Word document must have hyperlink and alt text).
'
'
Dim aWDoc As Document, shp As InlineShape, lnk As Hyperlink
Dim x As Integer, txt As String, url As String, quote As String, str As String

Set aWDoc = ActiveDocument
quote = """"
x = 1

For Each shp In aWDoc.InlineShapes
    url = InputBox("IMPORTANT!" & vbCrLf & "Enter server path for image #" & x & " (leave blank to adjust manually in HTML). Do upload the image to the server (this macro cannot do it for you).")
    txt = InputBox("IMPORTANT!" & vbCrLf & "Enter 'alt' text for image #" & x & ".")
    str = "<img src=" + quote + url + quote + " alt=" + quote + txt + quote + ">"
    shp.Select
    Selection.InsertBefore str
    shp.Delete
    x = x + 1
Next shp

For Each lnk In aWDoc.Hyperlinks
    url = lnk.Address
    txt = lnk.TextToDisplay
    str = "<a href=" + quote + url + quote + " target=" + quote + "_blank" + quote + ">" + txt + "</a>"
    lnk.TextToDisplay = str
Next lnk

aWDoc.Select
    Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
        .Text = "../"
        .Replacement.Text = ""
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

End Sub

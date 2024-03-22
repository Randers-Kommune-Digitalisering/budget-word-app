VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub indsætTabelbeskrivelser()

Dim comments As String

Dim commentsArrayS0() As String
Dim commentsArrayS1() As String
Dim nr() As String
Dim nr_s As String
Dim titel() As String
Dim titel_s As String
Dim beskrivelse() As String
Dim beskrivelse_s As String
Dim i As Integer

Dim t As Table

comments = ActiveDocument.BuiltInDocumentProperties("Comments")

comments = Left(comments, Len(comments) - 1)
comments = Right(comments, Len(comments) - 1)

commentsArrayS0 = Split(comments, "{")

i = 1
For Each t In ActiveDocument.Tables
    commentsArrayS0(i) = Replace(commentsArrayS0(i), "},", "")
    commentsArrayS0(i) = Replace(commentsArrayS0(i), "}", "")

    commentsArrayS1 = Split(commentsArrayS0(i), "," & Chr(34))

    nr = Split(commentsArrayS1(0), ":")
    nr_s = Replace(nr(1), Chr(34), "")
    titel = Split(commentsArrayS1(1), ":")
    titel_s = Replace(titel(1), Chr(34), "")
    beskrivelse = Split(commentsArrayS1(2), ":")
    beskrivelse_s = Replace(beskrivelse(1), Chr(34), "")

    t.Descr = beskrivelse_s
    t.Title = titel_s
    
    i = i + 1
Next

MsgBox "Tabeltitler og beskrivelser indsat"

End Sub


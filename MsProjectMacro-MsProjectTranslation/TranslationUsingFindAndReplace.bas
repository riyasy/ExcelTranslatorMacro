Attribute VB_Name = "Module2"
Sub TranslateFromJpToEn()
    Dim filepath As String
    Dim line As String
    Dim arrayOfElements() As String
    Dim linenumber As Integer
    Dim strLine As String


    filepath = "D:\macro-en-jp.txt"
    
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Type = 2
    objStream.Open
    objStream.LoadFromFile = filepath
    objStream.LineSeparator = 10
	OutlineShowAllTasks
    Do Until objStream.EOS
        strLine = objStream.ReadText(-2)
        arrayOfElements = Split(strLine, "|")
        ReplaceEx Field:="Name", Test:="contains exactly", Value:=arrayOfElements(0), Replacement:=arrayOfElements(1), ReplaceAll:=True, Next:=True, MatchCase:=False, SearchAllFields:=False
    Loop
End Sub

Sub TranslateFromEnToJp()
    Dim filepath As String
    Dim line As String
    Dim arrayOfElements() As String
    Dim linenumber As Integer
    Dim strLine As String


    filepath = "D:\macro-en-jp.txt"
    
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Type = 2
    objStream.Open
    objStream.LoadFromFile = filepath
    objStream.LineSeparator = 10
	OutlineShowAllTasks
    Do Until objStream.EOS
        strLine = objStream.ReadText(-2)
        arrayOfElements = Split(strLine, "|")
        ReplaceEx Field:="Name", Test:="contains exactly", Value:=arrayOfElements(1), Replacement:=arrayOfElements(0), ReplaceAll:=True, Next:=True, MatchCase:=False, SearchAllFields:=False
    Loop
End Sub



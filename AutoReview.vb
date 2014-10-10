Sub InsertComment(par As Paragraph, firstIndex As Long, length As Long, message As String)
'
' This macro inserts a comment in paragraph @param par
' with the scope starting at position @param firstIndex
' and of scope length of @param length.
' The content of the comment is given by the @param message.
'
	Dim myRange As Range
	Dim cmt As Comment
	
	Set myRange = par.Range
	
	' We want to select the specified range in the paragraph and store it in myRange.
	myRange.SetRange par.Range.Characters(firstIndex + 1).start, _
		par.Range.Characters(firstIndex + length).End
	
	' Then we attach a comment cmt to the range, with a diagnostic message.
	Set cmt = ActiveDocument.Comments.Add(myRange, message)
	
	With cmt ' some niceties
		.Author = "Auto Review" ' set the comment author's name
		.Initial = "Auto" ' set the label that appears in the bubble
	End With
End Sub



Sub findAndComment(pattern As String, message As String)
'
' This macro finds mistakes following the @param pattern
' and applies comments with the message @param message
'
	' Requires the "Microsoft VBScript Regular Expressions 5.5" library enabled in Tools > References...
	Dim regEx As New VBScript_RegExp_55.RegExp
	regEx.pattern = pattern
	regEx.IgnoreCase = True
	regEx.Global = True
	' set .Global to True to match all occurrences,
	' False matches only the first occurrence

	' auxiliary variables
	Dim matches As VBScript_RegExp_55.MatchCollection
	Dim str As String
	Dim line As Paragraph

		For Each line In ActiveDocument.paragraphs ' Iterate over all paragraphs in the document.
		str = line.Range.Text ' The string we are looking in is the text of the whole paragraph.
		
		If regEx.Test(str) Then ' First, check if there are any matches in the string at all.
			Set matches = regEx.Execute(str) ' If there are, gather all matches in a collection.
			
			For Each Match In matches
			' Iterate over the collection of matches.
			' Each Match is a range of text in the paragraph, where a potential misalignment is.
			
				InsertComment line, Match.firstIndex, Match.length, message
			Next
		End If
	Next
End Sub



Sub AutoReview()
'
' This macro tries to catch small tag misalignments,
' which are hard to catch exhaustively during a peer review,
' and which are common for inexperienced KEs
'

	' Requires the "Microsoft Scripting Runtime" library enabled in Tools > References...
	Dim fso As New FileSystemObject
	'array of values in each line
	Dim LineValues() As String
	'each new line read in from the text stream
	Dim ReadLine As String
	Dim ts As TextStream
	
	If Not fso.FileExists("C:\Users\" + uName + "\Desktop\AutoReviewPatterns.txt") Then
		MsgBox "Can't find the patterns file. Put the patterns file, ""AutoReviewPatterns.txt"", on your Desktop, and then try again."
	Else
		Set ts = fso.OpenTextFile("C:\Users\" + uName + "\Desktop\AutoReviewPatterns.txt")
		
		Do Until ts.AtEndOfStream
			ReadLine = ts.ReadLine
			LineValues = Split(ReadLine, Chr(9))
			If LineValues(2) = "Y" Then
				findAndComment LineValues(0), LineValues(1)
			End If
		Loop
	End If
End Sub



Function uName()
    Dim nameParts() As String
    nameParts = Split(Application.UserName, " ")
    
    uName = Left(nameParts(0), 1) + nameParts(1)
End Function
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

	' The pattern to catch:
	' a word (possibly followed by punctuation or quotes)
	' followed by a [p], [point], or [points] instruction without specifying what's being pointed at
	'
	findAndComment "[a-zA-Z]+\s*['""]?\s*[.,:;]?\s*['""]?\s*\[\s*p(oint(s)?)?\s*]", _
		"There's a word in front of this [point]. A [point]'s default argument can only be a number."
	
	' The pattern to catch:
	' a word (possibly followed by punctuation or quotes)
	' followed by a [show] instruction without specifying what's being pointed at
	'
	findAndComment "[a-zA-Z]+\s*['""]?\s*[.,:;]?\s*['""]?\s*\[\s*show\s*]", _
		"There's a word in front of this [show]. A [show]'s default argument can only be a number."
	
	' The pattern to catch:
	' a word (possibly followed by punctuation or quotes)
	' followed by a [i], [indicate], or [indicates] instruction without specifying what's being pointed at
	'
	findAndComment "[a-zA-Z]+\s*['""]?\s*[.,:;]?\s*['""]?\s*\[\s*i(ndicate(s)?)?\s*]", _
		"There's a word in front of this [indicate]. An [indicate]'s default argument can only be a number."
	
	' The pattern to catch:
	' a word (possibly followed by punctuation or quotes)
	' followed by a [w&t] instruction without specifying what's being pointed at
	'
	findAndComment "[a-zA-Z]+\s*['""]?\s*[.,:;]?\s*['""]?\s*\[\s*w\s*&\s*t\s*]", _
		"There's a word in front of this [w&t]. A [w&t]'s default argument can only be a number."
	
	' The pattern to catch:
	' an animator note splitting a number
	'
	findAndComment "\d\s*\[(?:(?!]).)+?]\s*\d", _
		"Either a punctuation is missing after a number or a number is split by the animator note."
	
	
	' The pattern to catch:
	' a punctuation mark right after an animator note
	'
	findAndComment "]\s*([.,:;])", _
		"There's a punctuation mark right after the animator note."
	
	' The pattern to catch:
	' a word in front of a [w&t], [show], [popup], or [pop-up] instruction
	' that doesn't match the first word being shown
	findAndComment "(\b\w+)\s*[,.:;]?\s*\[\s*(?:w\s*&\s*t|show|pop-?up)\s*[:\s]\s*(?:(?!\1)|(?=\1\w))\w+", _
		"Possible a/v sync issue. Voice doesn't match what appears on the screen."
End Sub
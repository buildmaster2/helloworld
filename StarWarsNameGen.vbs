'Star Wars Name Generator v2

'Does a comment count?
'Declare variables to erase any data from them.
Dim strSWName, strSWFirstName, strSWLastName, strSWHonorName, strMedication

'Set the title text for the message and input boxes.
strTitle = "Star Wars Name Generator v2"

'Star Wars First Name
'The rule: first 3 letters of first name + first 2 letters of last name.
strFirstName = InputBox("What is your first name?",strTitle)
	If strFirstName = "" Then Wscript.Quit
strLastName = InputBox ("What is your last name?",strTitle)
	If strLastName = "" Then Wscript.Quit
strSWFirstName = Left(strFirstName,3) & Left(strLastName,2)
'Set the capital letter.
strSWFirstName = funcCapitalize(strSWFirstName)

'Build Star Wars name to display in next message box.
strSWName = strSWFirstName

'Star Wars Last Name
'The rule: 1st 2 letters of mother's maiden name + 1st 3 letters of city you were born.
strMomMaiden = InputBox("Star Wars name: " & strSWName & vbCrLf & vbLf & _
	"Your mother's maiden name:",strTitle)
	If strMomMaiden = "" Then Wscript.Quit
strBirthCity = InputBox("Star Wars name: " & strSWName & vbCrLf & vbLf & _
	"The city in which you were born:",strTitle)
	If strBirthCity = "" Then Wscript.Quit
strSWLastName = Left(strMomMaiden,2) & Left(strBirthCity,3)
'Set the capital letter.
strSWLastName = funcCapitalize(strSWLastName)

'Build Star Wars name to display in next message box.
strSWName = strSWName & " " & strSWLastName

'Star Wars Honorific Name
'The rule: Last three letter of last name, reversed, + name of first car you drove/owned;
'          then add "of" and the name of the last medication you took.
strFirstCar = InputBox("Star Wars name: " & strSWName & vbCrLf & vbLf & _
	"What is the name of the first car you owned/drove?",strTitle)
	If strFirstCar = "" Then Wscript.Quit
'First, make first part of name.
strSWHonorName = StrReverse(Right(strLastName,3)) & strFirstCar
'Set the capital letter.
strSWHonorName = funcCapitalize(strSWHonorName)

'Build Star Wars name to display in next message box.
strSWName = strSWName & ", " & strSWHonorName

'Get medication name.
strMedication = InputBox("Star Wars name: " & strSWName & vbCrLf & vbLf & _
	"What is the name of the last medication you took?",strTitle)
	If strMedication = "" Then Wscript.Quit
'Set the capital letter.
strMedication = funcCapitalize(strMedication)

'Add the last piece of the honorific name.
strSWName = strSWName & " of " & strMedication

'Display the entire name.
MsgBox "Your Star Wars Universe name is:" & vbNewLine & _
	"""" & strSWName & """" & vbNewLine & _
	"Click OK to copy name to clipboard and end.", vbOKOnly, strTitle
'Copy Star Wars name to clipboard.
Set objShell = CreateObject("WScript.Shell")
objShell.Run "cmd /c echo " & strSWName & " | CLIP", 2

'Function to capitalize a string.
Function funcCapitalize(strString)
	funcCapitalize = UCase(Left(strString,1)) & LCase(Right(strString,Len(strString)-1))
End Function

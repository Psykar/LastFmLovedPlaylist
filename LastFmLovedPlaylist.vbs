Option Explicit
'==========================================================================
'
' MediaMonkey Script
'
' SCRIPTNAME: Last.fm Loved Tracks Playlist Creator
' DEVELOPMENT STARTED: 2009.02.17
  Dim Version : Version = "1.0"

' DESCRIPTION: Create a playlist containing loved tracks of a specific last.fm user
' FORUM THREAD: http://www.mediamonkey.com/forum/viewtopic.php?f=2&t=15663&start=15#p191962
' 
' INSTALL: Copy to Scripts directory and add the following to Scripts.ini 
'          Don't forget to remove comments (') and set the order appropriately
'
'
' [LastFmImport]
' FileName=LastFmLovedPlaylist.vbs
' ProcName=LastFmLovedPlaylist
' Order=7
' DisplayName=Last FM Loved Tracks Playlist Creator
' Description=Create a playlist containing loved tracks of a specific last.fm user
' Language=VBScript
' ScriptType=0 
'

Const ForReading = 1, ForWriting = 2, ForAppending = 8, Logging = False, Timeout = 100




Sub LastFmLovedPlaylist

	' Stats variables
	Dim Counter
	' Update logfile variables
	Dim fso, updatef,tmp

	' Last.fm username
	dim uname

	'XML result holder
	Dim XML


	' Which page of loved tracks we are processing
	Dim Page
	
	' String containing tracks not added / duplicate tracks
	Dim Not_Added, Duplicate_Tracks, Num_Dupes

	' Playlist for tracks to be added to
	Dim Playlist

	' Loop element
	Dim Ele
	

	Dim TrackTitle, ArtistName
	
	' temp list of tracks which match a loved track
	Dim list
	
	' temp holder for max number of plays of a track, so we can choose if we have duplicates.
	Dim Plays
	' temp holder for track to be added to the playlist
	Dim Add_me
	' Status Bar
	Dim StatusBar
	Set StatusBar = SDB.Progress
  
	StatusBar.Text = "Getting UserName"


	uname=InputBox("Enter your Last.fm username:")


	If uname = "" Then
		Exit Sub
	End If

	

	StatusBar.Text = "Loading Loved Tracks Tracks..."
	Set XML = LoadXML(uname,1)
	SDB.ProcessMessages


	If Not XML.getElementsByTagName("lfm").item(0).getAttribute("status") = "ok" Then
		SDB.MessageBox "Error" & VbCrLf & XML.getElementsByTagName("lfm").item(0).getElementsByTagName("error").item(0).text,mtInformation,Array(mbOK)
		Exit Sub
	End If
	'logme " ChartListXML appears to be OK, proceeding with loading each weeks data"



	

	Counter = 0
	Not_Added = ""
	Duplicate_Tracks = ""
	StatusBar.MaxValue = XML.getElementsByTagName("lfm").item(0).getElementsByTagName("lovedtracks").item(0).getAttribute("totalPages")

	Set Playlist = SDB.PlaylistByTitle("").CreateChildPlaylist("Loved Tracks (" & uname & ")")
	Playlist.Clear


	For Page = StatusBar.MaxValue To 1 Step -1
		
		Counter = Counter + 1
		StatusBar.Text = "Loading Loved Tracks... (Page " & Counter & " of " & StatusBar.MaxValue & ")"
		StatusBar.Increase
		If StatusBar.Terminate Then
		  Exit Sub
		End If

		logme " Page: " & Page
		Set XML = LoadXML(uname, Page)
		SDB.ProcessMessages


		If NOT (XML Is Nothing) Then
			If Not XML.getElementsByTagName("lfm").item(0).getAttribute("status") = "ok" Then
				SDB.MessageBox "Error" & VbCrLf &  XML.getElementsByTagName("lfm").item(0).getElementsByTagName("error").item(0).text,mtInformation,Array(mbOK)
				Exit Sub
			End If
			'logme "XML appears to be OK, proceeding"


		
			For Each Ele in XML.GetElementsByTagName("lfm").item(0).GetElementsByTagName("lovedtracks").item(0).GetElementsByTagName("track")

				TrackTitle = Ele.GetElementsByTagName("name").item(0).Text
				ArtistName = Ele.GetElementsByTagName("artist").item(0).GetElementsByTagName("name").item(0).Text
				

				Set list = SDB.Database.QuerySongs("Artist = '" & CorrectSt(ArtistName) & "' AND SongTitle = '" & CorrectSt(TrackTitle) & "'")

				If list.EOF Then
					Not_Added = Not_Added & VbCrLf & TrackTitle & " - " & ArtistName
				End If

				Plays = -1
				Num_Dupes = 0

				Do While Not list.EOF
					logme list.item.Title
					Num_Dupes = Num_Dupes + 1
					If list.item.Playcounter > Plays Then
						Set Add_me = list.item
						Plays = list.item.Playcounter
					End If
					list.next
				Loop
				If Plays >= 0 Then
					Playlist.AddTrack(Add_me)
				End If

				If Num_Dupes > 1 Then
					Duplicate_Tracks = Duplicate_Tracks & VbCrLf & TrackTitle & " - " & ArtistName
				End If
					
				

				SDB.ProcessMessages
			 
			Next
		Else
			Exit Sub
		End If
		SDB.ProcessMessages

	Next
	SDB.ProcessMessages
	
	If Not Not_Added = "" Then
		Not_Added = "The following tracks were not added because they are not in your database:" & Not_Added
	End If

	If Not Duplicate_Tracks = "" Then
		Duplicate_Tracks = VbCrLf & VbCrLf & "You also have duplicates of these tracks in your database" &_
						VbCrLf & "The one with the highest playcount was added to the playlist:" & VbCrLf & Duplicate_Tracks
		SDB.MessageBox Not_Added & Duplicate_Tracks,mtInformation,Array(mbOK)
	End If

End Sub

'**********************************************************


Function LoadXML(User,Page)
	'LoadXML accepts input string and mode, returns xmldoc of requested string and mode'
	'http://msdn2.microsoft.com/en-us/library/aa468547.aspx'
	logme ">> LoadXML: Begin with " & User & " & " & Page
	Dim xmlDoc, xmlURL, StatusBar, LoadXMLBar, StartTimer, http, strippedText
	StartTimer = Timer

	xmlURL = "http://ws.audioscrobbler.com/2.0/?method=user.getlovedtracks&user=" &_
		fixurl(user) & "&api_key=daadfc9c6e9b2c549527ccef4af19adb&limit=50&page=" &_
		Page

	logme ">> URL: " & xmlURL

	Set xmlDoc = CreateObject("MSXML2.DOMDocument.3.0")
	Set http = CreateObject("Microsoft.XmlHttp")
	
	http.open "GET",xmlURL,True
	http.send ""
	

	StartTimer = Timer
	'Wait for up to 3 seconds if we've not gotten the data yet
	  Do While http.readyState <> 4 And Int(Timer-StartTimer) < Timeout
		SDB.ProcessMessages
		SDB.Tools.Sleep 100
		SDB.ProcessMessages
	  Loop

	  If (http.readyState <> 4) Then
		SDB.MessageBox "HTTP request timed out. No tracks updated",mtInformation,Array(mbOK)
		Set LoadXML = Nothing
		Exit Function
	End If

	strippedText = stripInvalid(http.responseText)
	'MsgBox "Post Text: " & strippedText

	xmlDoc.async = True 
	xmlDoc.LoadXML(strippedText)

	StartTimer = Timer
	'Wait for up to 3 seconds if we've not gotten the data yet
	  Do While xmlDoc.readyState <> 4 And Int(Timer-StartTimer) < Timeout
		SDB.ProcessMessages
		SDB.Tools.Sleep 100
		SDB.ProcessMessages
	  Loop

	If (xmlDoc.parseError.errorCode <> 0) Then
		Dim myErr
		Set myErr = xmlDoc.parseError
		SDB.MessageBox "You have an error: " & myErr.reason,mtInformation,Array(mbOK)
		Set LoadXML = Nothing
	Else
		Dim currNode
		Set currNode = xmlDoc.documentElement.childNodes.Item(0)
	End If

	'logme " xmlDoc.Load: Waiting for Last.FM to return " & Mode & " of " & User
	SDB.ProcessMessages

	StartTimer = Timer
	Do While xmlDoc.readyState <> 4 And Int(Timer-StartTimer) < Timeout
		SDB.ProcessMessages
		SDB.Tools.Sleep 100
		SDB.ProcessMessages
	Loop



	'logme " xmlDoc: returned from loop in: " & (Timer - StartTimer)

	If xmlDoc.readyState = 4 and xmlDoc.parseError.errorCode = 0 Then 'all ok
		Set LoadXML = xmlDoc
		'msgbox("Last.FM query took: " & (timer-starttimer))
	Else
		'logme "Last.FM Query Failed @ " & Int(Timer-StartTimer) &	"ReadyState: " & xmlDoc.ReadyState & " URL: " & xmlURL
		SDB.MessageBox "Last.FM Query Failed",mtInformation,Array(mbOK)
		Set LoadXML = Nothing 
	End if

	'logme "<< LoadXML: Finished in --> " & Int(Timer-StartTimer)

End Function


'******************************************************************
'**************** Auxillary Functions *****************************
'******************************************************************

Sub logme(msg)
	'by psyXonova'
	If Logging Then
		'MsgBox "Yes!"
		Dim fso, logf
		On Error Resume Next
		Set fso = CreateObject("Scripting.FileSystemObject")
		'msgbox("logging: " & msg)
		Set logf = fso.OpenTextFile(Script.ScriptPath&".log",ForAppending,True)
		logf.WriteLine Now() & ": " & msg
		Set fso = Nothing
		Set logf = Nothing
	End If
End Sub



Function CorrectSt(inString)
' 	'logme ">> CorrectSt() has started with parameters: " & inString
	CorrectSt = Replace(inString, "'", "''")
	CorrectSt = Replace(CorrectSt, """", """""")
' 	'logme "<< CorrectSt() will return: " & CorrectSt & " and exit"
End Function


Function fixurl(sRawURL)
	' Original psyxonova improved by trixmoto
	'logme ">> fixurl() entered with: " & sRawURL
	Const sValidChars = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz\/!&:."
	sRawURL = Replace(sRawURL,"+","%2B")

	If UCase(Right(sRawURL,6)) = " (THE)" Then
		sRawURL = "The "&Left(sRawURL,Len(sRawURL)-6)
	End If
	If UCase(Right(sRawURL,5)) = ", THE" Then
		sRawURL = "The "&Left(sRawURL,Len(sRawURL)-5)
	End If

	If Len(sRawURL) > 0 Then
		Dim i : i = 1
		Do While i < Len(sRawURL)+1
			Dim s : s = Mid(sRawURL,i,1)
			If InStr(1,sValidChars,s,0) = 0 Then
				Dim d : d = Asc(s)
				If d = 32 Or d > 2047 Then
					s = "+"
				Else
					If d < 128 Then
						s = Hex(d)
					Else
						s = DecToUtf(d)
					End If
					s = "%" & s
				End If
			Else
				Select Case s
					Case "&"
						s = "%2526"
					Case "/"
						s = "%252F"
					Case "\"
						s = "%5C"
					Case ":"
						s = "%3A"
				End Select
			End If
			fixurl = fixurl&s
			i = i + 1
		SDB.ProcessMessages
    Loop
	End If
	'logme "<< fixurl will return with: " & fixurl
End Function




Function stripInvalid(str)
	Dim re, newStr, i

	Set re = new regexp
	Const invalidChars = "[\0\1\2\3\4\5\6\7\10\13\14\16\17\20\21\22\23\24\25\26\27\30\31\32\33\34\35\36\37]"
	newStr = str
	' Invalid: 0<=i<=8 or 11<=i<=12 or 14<=i<=31
	' Octal pattern of invalid chars
	re.Pattern = invalidChars
	Do While re.Test(newStr) = True
		newStr = re.Replace(newStr,"")
		'logme "==============Invalid character on this one!!???"
	Loop



	'logme "New text: " & VbCrLf & newStr & VbCrLf & "============================"
	stripInvalid = newStr
End Function 



'************************************************************'

' Thanks to trixmoto for this function
Sub Install()
	Dim inip : inip = SDB.ApplicationPath&"Scripts\Scripts.ini"
	Dim inif : Set inif = SDB.Tools.IniFileByPath(inip)
	If Not (inif Is Nothing) Then
		inif.StringValue("LastFmImport","Filename") = "LastFmLovedPlaylist.vbs"
		inif.StringValue("LastFmImport","Procname") = "LastFmLovedPlaylist"
		inif.StringValue("LastFmImport","Order") = "7"
		inif.StringValue("LastFmImport","DisplayName") = "Last FM Loved Tracks Playlist Creator"
		inif.StringValue("LastFmImport","Description") = "Create a playlist containing loved tracks of a specific last.fm user"
		inif.StringValue("LastFmImport","Language") = "VBScript"
		inif.StringValue("LastFmImport","ScriptType") = "0"
		SDB.RefreshScriptItems
	End If
End Sub

Option Explicit

Const GBOU_LINK = "https://komsomol.minobr63.ru/"
Const GIT_LINK = "https://github.com/KSOSH/TimeTableWord2Pdf/"
Const PROJECTSOFT_LINK = "https://projectsoft.ru/"

Const PDF = 17
Const cPrefixTitle = " КЛАССА НА "
Const windowW = 900
Const windowH = 650

Dim WShell: Set WShell = CreateObject("WScript.Shell")
Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objDlg: Set objDlg = CreateObject("Shell.Application")
Dim regExc: Set regExc = CreateObject("VBScript.RegExp")
Dim ReadIni: Set ReadIni = CreateObject("Scripting.Dictionary")	
Dim objFile

Dim objPlayer
Dim byff

Dim tGboy
tGboy = "ГБОУ СОШ пос. Комсомольский м. р. Кинельский Самарской обл."
Dim startPath
startPath = WShell.SpecialFolders.Item("Desktop") & "\Дистанционное"

Dim srvTimeTable
Dim winTimeTable
Dim strTimeTable

Dim HELP: Set HELP = document.getElementById("HELP")
Dim output: Set output = document.getElementById("output")
Dim btnSelectPath: Set btnSelectPath = document.getElementById("btnSelectPath")
Dim folderPath: Set folderPath = document.getElementById("folderPath")
Dim folderPathText: Set folderPathText = document.getElementById("folderPathText")
Dim OU: Set OU = document.getElementById("OU")
Dim gb_link: Set gb_link = document.getElementById("GBOU")
Dim ps_link: Set ps_link = document.getElementById("ProjectSoft")
Dim ProgressLine: Set ProgressLine = document.getElementById("progress_line")
Dim ProgressVal: Set ProgressVal = document.getElementById("progress_value")
Dim SRV: Set SRV = document.getElementById("SRV")
Dim btnConvert: Set btnConvert = document.getElementById("btnConvert")
Dim btnHelp: Set btnHelp = document.getElementById("btnHelp")
Dim btnClose: Set btnClose = document.getElementById("btnClose")
Dim BGSOUND: Set BGSOUND = document.getElementById("BGSOUND")
Dim SND_1: Set SND_1 = document.getElementById("SND_1")
Dim SND_2: Set SND_2 = document.getElementById("SND_2")

window.resizeTo windowW, windowH
window.moveTo (window.screen.availWidth - windowW) / 2, (window.screen.availHeight - windowH) / 2
ReadIni.Add "GBOU", tGboy
ReadIni.Add "SRV", "assets/files/0000/do"

ReadIniFile "config.cfg"
Reset
CheckData
Dim HWND
HWND = &H0

Sub SaveSounds()
	Dim fileName, winPath, id, val, inp
	Dim RE: Set RE = CreateObject("VBScript.RegExp")
	RE.Pattern = "data\:audio\/x-wav;base64,"
	winPath = WShell.ExpandEnvironmentStrings("%AppData%") & "\TimeTableWord2PDF\Media\"
	CreateFolderRecursive(winPath)
	For id = 1 to 2
		fileName = winPath & "snd-000" & CStr(id) & ".wav"
		Set Inp = document.getElementById("SND_" & CStr(id))
		If Not objFSO.FileExists(fileName) Then
			val = Inp.innerText
			val = RE.Replace(val, "")
			val = Base64Decode(val)
			With objFSO.createTextFile(fileName)
				.Write(val)
				.Close
			End With
		End If
		Inp.innerText = fileName
		Set Inp = Nothing
	Next
End Sub

Sub SaveImages()
	Dim fileName, winPath, id, val, inp, fln
	Dim RE: Set RE = CreateObject("VBScript.RegExp")
	RE.Pattern = "data\:image\/jpeg;base64,"
	CreateFolderRecursive(winPath)
	winPath = WShell.ExpandEnvironmentStrings("%AppData%") & "\TimeTableWord2PDF\Images\"
	For id = 0 to 6
		Set Inp = document.getElementById("IMG_" & CStr(id))
		fileName = winPath & RIGHT(String(4, "0") & CStr(id), 4) & ".jpg"
		If Not objFSO.FileExists(fileName) Then
			val = Inp.SRC
			val = Base64Decode(RE.Replace(val, ""))
			With objFSO.createTextFile(fileName)
				.Write(val)
				.Close
			End With
		End If
		Inp.SRC = fileName
	Next
End Sub

SaveSounds
saveImages

' Создадим папку "Дистанционное". Данная папка должна всегда существовать на рабочем столе. 
' Будем работать только в ней, тем самым не засоряем систему. 
' В данную папку нужно размещать директории с папками расписаний.
' У меня это выглядит следующим образом
' C:\Users\Администратор\Desktop\Дистанционное\Расписание уроков\21.10.2019
' C:\Users\Администратор\Desktop\Дистанционное\Расписание уроков\22.10.2019
' C:\Users\Администратор\Desktop\Дистанционное\Расписание уроков\23.10.2019
' и т. д.
' C:\Users\Администратор\Desktop\Дистанционное\Внеурочная деятельность\21.10.2019
' C:\Users\Администратор\Desktop\Дистанционное\Внеурочная деятельность\22.10.2019
' C:\Users\Администратор\Desktop\Дистанционное\Внеурочная деятельность\23.10.2019
' и т. д.
If Not objFSO.FolderExists(startPath) Then
	objFSO.CreateFolder(startPath)
End If

' Функция выбора директории
Function openBrowserDlg
	Dim objFolder, Result
	Result = ""
	' Запускаем диалог выбора директории
	' 512 - убрать кнопку "Создать папку". 1 - Выбирать только папки файловой системы. 512 + 1 = 513
	' 16 - установить начальную дирректорию "Рабочий стол" без отображения виртуальных папок.
	Set objFolder = objDlg.BrowseForFolder (HWND, strTimeTable & "." & vbCrlf & "Формат имени каталога - 	dd.mm.YYYY", 513, startPath) ' 0'
	' Если objFolder объект Folder
	If (Not objFolder Is Nothing) Then
		' Возвращаем путь до выбранной директории
		If StrComp(LCase(objFolder.Self.Path), LCase(startPath)) = 0  Then
			' Возвращаем пустую строку
			Result = ""
			MsgBox "Директорию """ & startPath & """ использовать нельзя.", vbExclamation, "Ошибка"
		Else
			Result = objFolder.Self.Path
			PlaySound SND_2
		End If
	End If
	openBrowserDlg = Result
End Function

Function CheckType
	' TypeTimeTable
	If TypeTimeTable(0).Checked Then
		winTimeTable = "\rs\"
		srvTimeTable = "/rs/"
		strTimeTable = "РАСПИСАНИЕ ЗАНЯТИЙ"
	Else
		winTimeTable = "\vd\"
		srvTimeTable = "/vd/"
		strTimeTable = "РАСПИСАНИЕ ЗАНЯТИЙ ВНЕУРОЧНОЙ ДЕЯТЕЛЬНОСТИ"
	End If
End Function

Function CheckFiles(path)
	Dim Result, oFile
	Result = False
	For Each oFile In objFSO.GetFolder(path).Files
		If StrComp(objFSO.GetExtensionName(oFile.Name), "docx", vbTextCompare) = 0 Or StrComp(objFSO.GetExtensionName(oFile.Name), "doc", vbTextCompare) = 0 Then
			regExc.Global = False
			regExc.Pattern = "^~\$"
			If Not regExc.Test(oFile.Name) Then
				Result = True
				Exit For
			End If
		End If
	Next
	'If Not oFile = Nothing Then
	'	oFile = Nothing
	'End If
	CheckFiles = Result
End Function

Function CreateFolderRecursive(FullPath)
	Dim arr, dir, path
	Dim oFs
	
	Set oFs = CreateObject("Scripting.FileSystemObject")
	arr = split(FullPath, "\")
	path = ""
	For Each dir In arr
		If path <> "" Then path = path & "\"
		path = path & dir
		If oFs.FolderExists(path) = False Then oFs.CreateFolder(path)
	Next
	Set oFs = Nothing
End Function

Function replaceSpace(fName)
	Dim str
	regExc.Global = True
	regExc.Pattern = "\s+"
	str = regExc.Replace(fName, " ")
	replaceSpace = str
End Function

Function removeSpace(fName)
	Dim str
	regExc.Global = True
	regExc.Pattern = "\s+"
	str = regExc.Replace(fName, "")
	removeSpace = str
End Function

Sub Reset
	RS.Checked = True
	ProgressLine.style.width = "0%"
	ProgressVal.innerText = "0%"
	'OU.Value = "ГБОУ СОШ пос. Комсомольский м. р. Кинельский Самарской обл."
	'SRV.Value = "assets/files/0000/do"
	folderPath.Value = ""
	folderPath.Title = ""
	folderPathText.innerHtml = "&nbsp;"
	folderPathText.Title = ""
	btnConvert.Disabled = True
	CheckType
End Sub

Sub CheckData
	If folderPath.Value = "" Then
		btnConvert.Disabled = True
	Else
		If CheckFiles(folderPath.Value) = True Then
			btnConvert.Disabled = False
		Else
			btnConvert.Disabled = True
		End If
	End If
End Sub

Sub fnShellParentVB(path)
	PlaySound SND_2
	dim objShell
	dim ssfWINDOWS
	set objShell = CreateObject("shell.application")
	objShell.Explore(path)
	set objShell = nothing
End Sub

Sub btnSelectPath_OnClick()
	PlaySound SND_2
	Dim Result
	Result = openBrowserDlg
	If Not Result = "" Then
		' folderPath.value = Result
		folderPathText.innerHtml = Result
	Else
		folderPathText.innerHtml = "&nbsp;"
	End If
	folderPath.Value = Result
	folderPath.Title = Result
	folderPathText.Title = Result
	CheckData
End Sub

Sub DisabledApp
	btnConvert.Disabled = True
	btnSelectPath.Disabled = True
	RS.Disabled = True
	VD.Disabled = True
End Sub

Sub EnabledApp
	CheckData
	btnSelectPath.Disabled = False
	RS.Disabled = False
	VD.Disabled = False
End Sub

Sub RS_OnChange()
	PlaySound SND_2
	CheckType
End Sub

Sub VD_OnChange()
	PlaySound SND_2
	CheckType
End Sub

Sub SRV_OnChange()
	'MsgBox SRV.Value
	If Not ReadIni.Exists("SRV") Then
		ReadIni.Add "SRV", SRV.Value
	End If
	ReadIni("SRV") = SRV.Value
	SaveSettings "config.cfg"
End Sub

Sub OU_OnChange()
	'MsgBox OU.Value
	If Not ReadIni.Exists("GBOU") Then
		ReadIni.Add "GBOU", OU.Value
	End If
	ReadIni("GBOU") = OU.Value
	SaveSettings "config.cfg"
End Sub

Sub GIT_OnClick()
	WShell.Run GIT_LINK
	PlaySound SND_2
End Sub

Sub ProjectSoft_OnClick()
	WShell.Run PROJECTSOFT_LINK
	PlaySound SND_2
End Sub

Sub GBOU_OnClick()
	WShell.Run GBOU_LINK
	PlaySound SND_2
End Sub

Sub btnHelp_OnClick()
	HELP.style.display = "block"
	PlaySound SND_2
End Sub

Sub btnClose_OnClick()
	HELP.style.display = "none"
	PlaySound SND_2
End Sub

Sub ImgClick(obj)
	WShell.Run obj.Src
End Sub

Sub btnConvert_OnClick()
	PlaySound SND_2
	Dim csvFile, strSourceFolder, outputDir, tFName, objWord, docTitle, Files, objDocument, customProp
	Dim rsDate, prop, fCount, prg, csvText, assetsFolder, count, current, out10, out5, fn
	strSourceFolder = folderPath.Value
	outputDir = startPath & "\PDF\" & objFSO.GetFolder(strSourceFolder).Name & winTimeTable
	assetsFolder = SRV.Value
	fCount = 0
	prg = "0%"
	out10 = ""
	out5 = ""
	Set objWord = Nothing
	If CheckFiles(strSourceFolder) = True Then
		CreateFolderRecursive(outputDir)
		DisabledApp
		Set Files = objFSO.GetFolder(strSourceFolder).Files
		Set csvFile = objFSO.CreateTextFile(outputDir & "csv.csv", True)
		count = Files.Count
		For Each objFile In Files
			DoEvents(0)
			If StrComp(objFSO.GetExtensionName(objFile.Name), "docx", vbTextCompare) = 0 Or StrComp(objFSO.GetExtensionName(objFile.Name), "doc", vbTextCompare) = 0 Then
				' Проверяем имя файла. Если имя файла начинается с ~$ то он временный
				regExc.Global = False
				regExc.Pattern = "^~\$"
				If Not regExc.Test(objFile.Name) Then
					DoEvents(0)
					fCount = fCount + 1
					current = CStr(Round((fCount * 100) / (count))) & "%"
					ProgressLine.style.width = current
					ProgressVal.innerText = current
					tFName = objFile.Name
					output.innerText = "Конвертирование: " + tFName
					DoEvents(0)
					' Запускаем Word если он ещё не запущен
					If objWord Is Nothing Then
						Set objWord = CreateObject("Word.Application")
					End If
					' Пустой заголовок
					docTitle = ""
					' Открываем документ
					DoEvents(0)
					Set objDocument = objWord.Documents.Open(objFile.Path)
					' Получаем объект свойст документа
					Set customProp = objDocument.BuiltinDocumentProperties
					' Получаем дату
					rsDate = removeSpace(objFSO.GetBaseName(removeSpace(strSourceFolder)) & "." & objFSO.GetExtensionName(strSourceFolder))
					' Собираем заголовок
					fn = objFSO.GetBaseName(UCase(removeSpace(objFile.Name)))
					docTitle = strTimeTable & " " & fn & cPrefixTitle & rsDate
					
					' Перебираем свойства документа
					For Each prop in customProp
						' Устанавливаем нужные свойства документа
						DoEvents(0)
						Select case prop.Name
							' Заголовок документа
							case "Title"
								prop.Value = docTitle & " " & OU.Value
							' Тема документа
							case "Subject"
								prop.Value = docTitle & " " & OU.Value
							' Автор документа
							case "Author"
								prop.Value = OU.Value
							' Компания
							case "Company"
								prop.Value = OU.Value
						End Select
					Next
					' Сохраняем документ как PDF. Транслит имени файла для сохранения
					' Так же сначало сохраниться сам документ перед конвертацией.
					objDocument.SaveAs2 objFSO.BuildPath(outputDir, Rus2Lat(removeSpace(objFSO.GetBaseName(objFile.Name))) & ".pdf"), PDF
					If StrComp(fn, "10") = 0 Or StrComp(fn, "11") = 0 Then
						' 10 - 11 классы отдельно
						out10 = out10 & """" & docTitle & """;""" & assetsFolder & "/" & rsDate & srvTimeTable & Rus2Lat(removeSpace(objFSO.GetBaseName(objFile.Name))) & ".pdf""" & vbCrlf
					Else
						' 5 - 9 классы отдельно
						out5 = out5 & """" & docTitle & """;""" & assetsFolder & "/" & rsDate & srvTimeTable & Rus2Lat(removeSpace(objFSO.GetBaseName(objFile.Name))) & ".pdf""" & vbCrlf
					End If
					' Закрываем документ
					objDocument.Close
					' Обнуляем переменную
					' Set objDocument = Nothing
				End If
			Else
				count = count - 1
			End If
		Next
		' Если Word запущен - закроем его
		If Not objWord Is Nothing Then
			objWord.Quit
		End If
		' Обнуляем переменную
		Set objWord = Nothing
		' Записываем данные в csv файл
		csvFile.Write(out5 & out10)
		' Закрываем csv файл
		csvFile.Close
		PlaySound SND_1
		If MsgBox("Открыть папку с результатом конвертирования?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes And count > 0 Then
			fnShellParentVB(outputDir)
			DoEvents(0)
		End If
	End If
	CheckData
	ProgressLine.style.width = "0%"
	ProgressVal.innerText = "0%"
	output.innerHtml = "&nbsp;"
	DoEvents(0)
	EnabledApp
End Sub

Sub DoEvents(sec)
	With CreateObject("Msxml2.ServerXMLHTTP")
		.Abort
		.Open "HEAD", "http://1.0.0.1", True
		.Send
		.WaitForResponse sec
	End With
End Sub

Sub ReadIniFile (FileName )
	ReadIni.RemoveAll
	Dim winPath, line, pos, FileStr, fln
	winPath = WShell.ExpandEnvironmentStrings("%AppData%") & "\TimeTableWord2PDF\"
	CreateFolderRecursive(winPath)
	fln = winPath & FileName
	Set FileStr = objFSO.OpenTextFile( fln, 1, True)
	While Not FileStr.AtEndOfStream
		line = FileStr.ReadLine()
		pos = InStr(line, "=")
			if pos = 0 then Continue
		ReadIni.Add Trim(Left(line, pos - 1)), Trim(Mid(line, pos + 1)) 
	Wend
	FileStr.Close
	If Not ReadIni.Exists("GBOU") Then
		ReadIni.Add "GBOU", "ГБОУ СОШ пос. Комсомольский м. р. Кинельский Самарской обл."
	End If
	If Not ReadIni.Exists("SRV") Then
		ReadIni.Add "SRV", "assets/files/0000/do"
	End If
	OU.Value = ReadIni("GBOU")
	SRV.Value = ReadIni("SRV")
	SaveSettings FileName
End Sub

Sub SaveSettings(strFile)
	Dim winPath, fln
	winPath = WShell.ExpandEnvironmentStrings("%AppData%") & "\TimeTableWord2PDF\"
	CreateFolderRecursive(winPath)
	fln = winPath & strFile
	With CreateObject("Scripting.FileSystemObject").CreateTextFile(fln, True)
		Dim k
		For Each k In ReadIni
			.WriteLine k & "=" & ReadIni(k)
		Next
		.Close
	End With
End Sub

Sub PlaySound(Input)
	BGSOUND.src = Input.innerText
	'If objFSO.FileExists(FileName) Then
		'Set objPlayer = CreateObject("Wmplayer.OCX.7")
		'With objPlayer  ' saves typing
		'	.settings.autoStart = True
		'	.settings.volume = 100  ' 0 - 100
		'	.settings.balance = 0  ' -100 to 100
		'	.settings.enableErrorDialogs = False
		'	.enableContextMenu = False
		'	.URL = FileName
		'End With
	'End If
End Sub

Function Rus2Lat(strRus)
	Dim i
	Dim strTemp
	Dim strLat
	strRus = LCase(strRus)
	For i = 1 To Len(strRus)
		strTemp = Mid(strRus, i, 1)			 
		Select Case strTemp
			Case "а"
				strLat = strLat & "a"
			Case "б"
				strLat = strLat & "b"
			Case "в"
				strLat = strLat & "v"
			Case "г"
				strLat = strLat & "g"
			Case "д"
				strLat = strLat & "d"
			Case "е"
				strLat = strLat & "e"
			Case "ё"
				strLat = strLat & "e"
			Case "ж"
				strLat = strLat & "zh"
			Case "з"
				strLat = strLat & "z"
			Case "и"
				strLat = strLat & "i"
			Case "й"
				strLat = strLat & "i"
			Case "к"
				strLat = strLat & "k"
			Case "л"
				strLat = strLat & "l"
			Case "м"
				strLat = strLat & "m"
			Case "н"
				strLat = strLat & "n"
			Case "о"
				strLat = strLat & "o"
			Case "п"
				strLat = strLat & "p"
			Case "р"
				strLat = strLat & "r"
			Case "с"
				strLat = strLat & "s"
			Case "т"
				strLat = strLat & "t"
			Case "у"
				strLat = strLat & "u"
			Case "ф"
				strLat = strLat & "f"
			Case "х"
				strLat = strLat & "kh"
			Case "ц"
				strLat = strLat & "ts"
			Case "ч"
				strLat = strLat & "ch"
			Case "ш"
				strLat = strLat & "sh"
			Case "щ"
				strLat = strLat & "sch"
			Case "ъ"
				strLat = strLat & ""
			Case "ы"
				strLat = strLat & "y"
			Case "ь"
				strLat = strLat & ""
			Case "э"
				strLat = strLat & "e"
			Case "ю"
				strLat = strLat & "yu"
			Case "я"
				strLat = strLat & "ya"
			case "«"
				strLat = strLat & ""
			case "»"
				strLat = strLat & ""
			case "№"
				strLat = strLat & "n"
			case " "
				strLat = strLat & " "
			Case Else
				strLat = strLat & strTemp
		End Select
	Next
	strLat = replaceSpace(strLat)
	Rus2Lat = strLat
End Function

Function Base64Decode(ByVal base64String)
	Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
	Dim dataLength, sOut, groupBegin
	base64String = Replace(base64String, """", "")
	base64String = Replace(base64String, vbCrLf, "")
	base64String = Replace(base64String, vbTab, "")
	base64String = Replace(base64String, " ", "")
	dataLength = Len(base64String)
	If dataLength Mod 4 <> 0 Then
		Err.Raise 1, "Base64Decode", "Bad Base64 string."
		Exit Function
	End If
	For groupBegin = 1 To dataLength Step 4
		Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
		numDataBytes = 3
		nGroup = 0
		For CharCounter = 0 To 3
			thisChar = Mid(base64String, groupBegin + CharCounter, 1)
			If thisChar = "=" Then
				numDataBytes = numDataBytes - 1
				thisData = 0
			Else
				thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
			End If
			If thisData = -1 Then
				Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
				Exit Function
			End If
			nGroup = 64 * nGroup + thisData
		Next
		nGroup = Hex(nGroup)
		nGroup = String(6 - Len(nGroup), "0") & nGroup
		pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
			Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
			Chr(CByte("&H" & Mid(nGroup, 5, 2)))
		sOut = sOut & Left(pOut, numDataBytes)
	Next
	Base64Decode = sOut
End Function
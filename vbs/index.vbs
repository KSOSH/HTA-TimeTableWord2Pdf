Option Explicit

Const PDF = 17
Const cPrefixTitle = " КЛАССА НА "
Const windowW = 1024
Const windowH = 740

Dim WShell: Set WShell = CreateObject("WScript.Shell")
Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objDlg: Set objDlg = CreateObject("Shell.Application")
Dim regExc: Set regExc = CreateObject("VBScript.RegExp")
Dim ReadIni: Set ReadIni = CreateObject("Scripting.Dictionary")	
Dim objFile

Dim objPlayer
Dim sndFile
sndFile =  WShell.ExpandEnvironmentStrings("%windir%") & "\Media\chimes.wav"

Dim tGboy
tGboy = "ГБОУ СОШ пос. Комсомольский м. р. Кинельский Самарской обл."
Dim startPath
startPath = WShell.SpecialFolders.Item("Desktop") & "\Дистанционное"
Dim tmpPath
tmpPath = WShell.SpecialFolders.Item("Temp") & "\Dist"

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

window.resizeTo windowW, windowH
window.moveTo (window.screen.availWidth - windowW) / 2, (window.screen.availHeight - windowH) / 2
ReadIni.Add "GBOU", tGboy
ReadIni.Add "SRV", "assets/files/0000/do"

ReadIniFile "config.cfg"
Reset
CheckData

'MsgBox TypeName(Window)
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
'CreateFolderRecursive(tmpPath)
' Функция выбора директории
Function openBrowserDlg
	Dim objFolder, Result
	Result = ""
	' Запускаем диалог выбора директории
	' 512 - убрать кнопку "Создать папку". 1 - Выбирать только папки файловой системы. 512 + 1 = 513
	' 16 - установить начальную дирректорию "Рабочий стол" без отображения виртуальных папок.
	Set objFolder = objDlg.BrowseForFolder (0, strTimeTable & "." & vbCrlf & "Формат имени каталога - 	dd.mm.YYYY", 513, startPath) ' 0'
	' Если objFolder объект Folder
	If (Not objFolder Is Nothing) Then
		' Возвращаем путь до выбранной директории
		If StrComp(LCase(objFolder.Self.Path), LCase(startPath)) = 0  Then
			' Возвращаем пустую строку
			Result = ""
			MsgBox "Директорию """ & startPath & """ использовать нельзя.", vbExclamation, "Ошибка"
		Else
			Result = objFolder.Self.Path
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

Sub btnSelectPath_OnClick()
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

Sub Window_OnResize()
	MsgBox "Resize"
End Sub

Sub RS_OnChange()
	CheckType
End Sub

Sub VD_OnChange()
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
	WShell.Run "https://github.com/KSOSH/TimeTableWord2Pdf"
End Sub

Sub ProjectSoft_OnClick()
	WShell.Run "https://projectsoft.ru/"
End Sub

Sub GBOU_OnClick()
	WShell.Run "https://komsomol.minobr63.ru/"
End Sub

Sub btnHelp_OnClick()
	HELP.style.display = "block"
End Sub

Sub btnClose_OnClick()
	HELP.style.display = "none"
End Sub

Sub ImgClick(obj)
	'WShell.Run obj.Src
	'window.open(obj.Src)
End Sub

Sub btnConvert_OnClick()
	Dim csvFile, strSourceFolder, outputDir, tFName, objWord, docTitle, Files, objDocument, customProp
	Dim rsDate, prop, fCount, prg, csvText, assetsFolder, count, current
	strSourceFolder = folderPath.Value
	outputDir = startPath & "\PDF\" & objFSO.GetFolder(strSourceFolder).Name & winTimeTable
	assetsFolder = SRV.Value
	fCount = 0
	prg = "0%"
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
					tFName = objFile.Name
					' Запускаем Word если он ещё не запущен
					If objWord Is Nothing Then
						Set objWord = CreateObject("Word.Application")
					End If
					output.innerText = tFName
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
					docTitle = strTimeTable & " " & objFSO.GetBaseName(UCase(removeSpace(objFile.Name))) & cPrefixTitle & rsDate
					
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
					' Записываем данные в csv файл
					csvText = """" & docTitle & """;""" & assetsFolder & "/" & rsDate & srvTimeTable & Rus2Lat(removeSpace(objFSO.GetBaseName(objFile.Name))) & ".pdf"""
					csvFile.WriteLine(csvText)
					' Закрываем документ
					objDocument.Close
					' Обнуляем переменную
					' Set objDocument = Nothing
					fCount = fCount + 1
					current = CStr(CInt((fCount * 100) / (count - 1))) & "%"
					ProgressLine.style.width = current
					ProgressVal.innerText = current
					DoEvents(0)
				End If
			End If
		Next
		' Если Word запущен - закроем его
		If Not objWord Is Nothing Then
			objWord.Quit
		End If
		' Обнуляем переменную
		Set objWord = Nothing
		' Закрываем csv файл
		csvFile.Close
		PlaySound sndFile
	End If
	CheckData
	ProgressLine.style.width = "0%"
	ProgressVal.innerText = "0%"
	output.innerHtml = "&nbsp;"
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
	Dim FileStr: Set FileStr = objFSO.OpenTextFile( FileName, 1, True)
	Dim line, pos
	While Not FileStr.AtEndOfStream
		line = FileStr.ReadLine()
		pos = InStr(line, "=")
			if pos = 0 then Continue
		ReadIni.Add Trim(Left(line, pos - 1)), Trim(Mid(line, pos + 1)) 
	Wend
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
	With CreateObject("Scripting.FileSystemObject").CreateTextFile(strFile, True)
		Dim k
		For Each k In ReadIni
			.WriteLine k & "=" & ReadIni(k)
		Next
	End With
End Sub

Sub PlaySound(FileName)
	If objFSO.FileExists(FileName) Then
		Set objPlayer = CreateObject("Wmplayer.OCX.7")
		With objPlayer  ' saves typing
			.settings.autoStart = True
			.settings.volume = 100  ' 0 - 100
			.settings.balance = 0  ' -100 to 100
			.settings.enableErrorDialogs = False
			.enableContextMenu = False
			.URL = FileName
		End With
	End If
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

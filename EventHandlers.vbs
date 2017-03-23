
Sub OnAcceptMessage(oClient, oMessage)	
	'==========Если письмо пришло===============
	if (InStr(oClient.Username, "@kolaer.ru") < 1 ) then
	
		'==========Фильтр разрешенных Доменов===============
		Dim arrayFilterDomens
		arrayFilterDomens = Array("@gmail.ru")
		
		'==============true - использовать фильт разрешенных доменов=============
		useDomenFilter = true
		
		'==========Фильтр исключений===============
		Dim arrayFilterExeptions
		arrayFilterExeptions = Array("danilov0x33@gmail.com")
		
		'==============true - использовать фильт исключений для наших email=============
		useFilterExeptions = true
		
		if (useFilterExeptions = true) then
			For w = 0 to UBound(arrayFilterExeptions) + 1
				if (InStr(oMessage.To, arrayFilterExeptions(w)) > 0 ) then
					Exit Sub
				end if
			
				if w = UBound(arrayFilterExeptions) then
					exit for
				end if
			Next
		end if

		isFilter = false
		If(useDomenFilter = true) Then 		
			For d = 0 to UBound(arrayFilterDomens) + 1
				if (InStr(oMessage.FromAddress, arrayFilterDomens(d)) > 0 ) then		
					isFilter = true
					Exit For
				end if
				
				if d = UBound(arrayFilterDomens) then
					exit for
				end if
			Next
			'Если письмо пришло от домена не в фильтре, значит запускаем скан на наличие плохих файлов
			if (isFilter = false) then
				ReadindEmail oClient, oMessage
			end if
		Else
			ReadindEmail oClient, oMessage
		End If	
	end if	
End Sub

'Чтение письма
Function ReadindEmail(oClient, oMessage)
	'===========Фильтр расширений файлов=============
	Dim arrayExtendsFile
	arrayExtendsFile = Array(".jar", ".bat", ".js", ".exe", ".cmd", ".com")
	
	'============true - Если нужно удалять все файлы. Иначе использовать фильтр===============
	blockAllFile = true
	
	'=============true - Если нужно сохранить удаленные/отфильтрованные файлы=================
	saveFileToDit = true

	'Если блокировать не все файлы...
	Dim path
	If(blockAllFile = false) Then 
		Dim isSave 
		isSave = false	
		
		'...иначе перебираем все файлы...
		For i = 0 to oMessage.Attachments.Count - 1
			For j = 0 to UBound(arrayExtendsFile) + 1
				'и проверяем на наличие расширения в фильтре
				if (InStr(oMessage.Attachments.item(i).Filename,arrayExtendsFile(j)) > 0 ) then
					if(saveFileToDit = true) then
						path = SaveFileToDir(oMessage, oMessage.Attachments.item(i))
						isSave = true
					end if
					oMessage.Attachments.item(i).Delete
					oMessage.Save()
					Exit For
				end if
			Next
			if i = oMessage.Attachments.Count then
				Exit For
			end if
		Next
		if(isSave = true) then
			oMessage.HTMLBody = "<h1><a href=""" + path + """ >Файлы в письме!</a></h1>" & oMessage.HTMLBody
			oMessage.Save()
		end if	
		ReadindEmail = true
	Else
		if(saveFileToDit = true) then
			isSave = false
			
			path = ""
			
			For i = 0 to oMessage.Attachments.Count - 1
				path = SaveFileToDir(oMessage, oMessage.Attachments.item(i))
				isSave = true
				
				if i = oMessage.Attachments.Count  then
					Exit For
				end if
			next
			
			if(isSave = true) then
				oMessage.HTMLBody = "<h1><a href=""" + path + """ >Файлы в письме!</a></h1>" & oMessage.HTMLBody
			end if
			
		end if

		oMessage.Attachments.Clear()
		oMessage.Save()
		
	End If
End Function

'Сохранить файл. userName - Имя пользователя; file - файл для сохранения
Function SaveFileToDir(oMessage, file)
	'Путь для сохранения файлов
	const pathToDataDir = "D:\email"
	'Путь для сохранения файлов + папка с именем пользователя
	fileSaveToDir = pathToDataDir + "\" + oMessage.To
	
	'Создание не сужествующих папок по пути в - fileSaveToDir
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	if(Not(objFSO.FolderExists(fileSaveToDir))) then
		objFSO.CreateFolder(fileSaveToDir)
	end if
	
	'Путь для сохранения файлов + папка с именем пользователя
	fileSaveToDirWithFrom = fileSaveToDir + "\" + oMessage.FromAddress
	
	if(Not(objFSO.FolderExists(fileSaveToDirWithFrom))) then
		objFSO.CreateFolder(fileSaveToDirWithFrom)
	end if

	dim my_date
	my_date = date
	
	'Путь для сохранения файлов + папка с именем пользователя + дата
	fileSaveToDirWithFromAndDate = fileSaveToDirWithFrom + "\" + CStr(my_date)

	if(Not(objFSO.FolderExists(fileSaveToDirWithFromAndDate))) then
		objFSO.CreateFolder(fileSaveToDirWithFromAndDate)
	end if

	file.SaveAs(fileSaveToDirWithFromAndDate + "\" + file.Filename)

	Dim arrayBlockFileInZip
	arrayBlockFileInZip = Array("Сценарий Windows", "Файл сценария JScript", "Пакетный файл Windows")

	Dim objSA, objSource, objZip, removeZip, item
	Set objSA = CreateObject("Shell.Application")
	Set objZip = objSA.NameSpace(fileSaveToDirWithFromAndDate + "\" + file.Filename)
	Set objSource = objZip.Items()
	
	removeZip = false
	
	For Each item in objSource
		'EventLog.Write(objZip.GetDetailsOf(item, 1))
		For c = 0 to UBound(arrayBlockFileInZip) + 1
			if (InStr(objZip.GetDetailsOf(item, 1),arrayBlockFileInZip(c)) > 0) then
				removeZip = true
				Exit For
			end if
			
			if c = UBound(arrayBlockFileInZip) then
				Exit For
			end if
		Next 
		
		if removeZip = true then
			Exit For
		end if
	Next 
	
	if removeZip = true then 
		EventLog.Write("Внутри плохие файлы:" + file.Filename)
		Set fso = CreateObject("Scripting.FileSystemObject")
		fso.DeleteFile fileSaveToDirWithFromAndDate + "\" + file.Filename
	end if
	
	SaveFileToDir = Replace(fileSaveToDirWithFromAndDate, "D:\email", "file:\\\\\D:\\email") + "\"
End Function

'    Sub OnClientConnect(oClient)
'    End Sub

'   Sub OnSMTPData(oClient, oMessage)
'   End Sub

'   Sub OnDeliveryStart(oMessage)
'   End Sub

'   Sub OnDeliverMessage(oMessage)
'   End Sub

'   Sub OnBackupFailed(sReason)
'   End Sub

'   Sub OnBackupCompleted()
'   End Sub

'   Sub OnError(iSeverity, iCode, sSource, sDescription)
'   End Sub

'   Sub OnDeliveryFailed(oMessage, sRecipient, sErrorMessage)
'   End Sub

'   Sub OnExternalAccountDownload(oFetchAccount, oMessage, sRemoteUID)
'   End Sub
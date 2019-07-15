'********************************************************************************************************
'	Скрипт просто удаляет установленные службы УТМ
'	Халиман Андрей , Рыбинск 2019 год . x00502@gmail.com
'
'********************************************************************************************************

'удаляет найденный сервис утм
'@path путь из реестра типа: E:\UTM\8081\monitoring\bin
'@tt имя службы , а именно одно из: Transport-Monitoring,Transport,Transport-Updater
sub deleteService(ByRef path,ByRef tt)

	A = Split(path, " //")
	I = A(0)
	WshShell.Run I&" //SS//"&tt
	WshShell.Run I&" //DS//"&tt

	if err.number <> 0 then
		
		MsgBox "Ошибка удаления службы "&tt
		WScript.Quit
		
	end if	

end sub 'deleteService

'*****Прочитаем из реестра адреса по которым зарегистрированы службы утм
Set WshShell = CreateObject("WScript.Shell")
Transport 				= WshShell.RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Transport\ImagePath")
TransportMonitoring 	= WshShell.RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Transport-Monitoring\ImagePath")
TransportUpdater 		= WshShell.RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Transport-Updater\ImagePath")

if err.number <> 0 then
	
	'***ничего нет в реестре , можно просто установить нужные сервисы
	MsgBox "Ошибка чтения реестра,возможно службы уже удалены"
	'WScript.Quit
else
	
	'**	в реестре что то есть , удаляем службы
	deleteService Transport,"Transport-Monitoring"
	deleteService TransportMonitoring,"Transport"
	deleteService TransportUpdater,"Transport-Updater"
	
end if	
'on error resume next
'********************************************************************************************************
'	������ ��������� �������� � ����������� ��� �� ����� ����� ���������������
'	������� ������ , ������� 2019 ��� . x00502@gmail.com
'	1. ��������� ��� � ����� �������� c:/UTM
'	2. ��������� ����� \transporter\transportDB � ����� �����
'	3. �������� ������ ����� UTM �������� 1,2,3 ���
'	4. ��������� ��� ����� � ����� 1, � ��� �� 2 � ��� �� 3 (����� 1 2 3 ���������� �� �����)
'	5. ������� ���� 1 � �������� ���� ������ �� ����� UTM , �������� ���� � ��� ������ 1 �������� ������� ������ � �������� ���� 1
'	6. ������� ���� 2 � �������� ���� ������ �� ����� UTM , �������� ���� � ��� ������ 2 �������� ������� ������ � �������� ���� 1
'	7. ��� �� ��������
'	8. � ������� myArray ������� ��� ������� ����� ������ �����. �������� �� ����� � ���� 8081 - ��� ����� �����
'	������� � ��� ��� ��� ������� ������� �� ����������������� ������ ��� �� ���������� ����������� 
'	� ��������� ������ ����� . ����� ���� �� ������� ������ ������. ����� �� ������� �� ����� - ����� ��������.
'
'********************************************************************************************************

Dim myArray,inputName
Dim WshShell

myArray = Array("1-�������", "2-������", "3-��������")
Set WshShell = CreateObject("WScript.Shell")

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'������ ���� � ����� , ��� ������ �����. �� ����� ������� ��������
function getVbsPath()

	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set F = FSO.GetFile(Wscript.ScriptFullName)

	getVbsPath = FSO.GetParentFolderName(F)

end function 'getVbsPath

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'������� ��������� ������ ���
'@path ���� �� ������� ����: E:\UTM\8081\monitoring\bin
'@tt ��� ������ , � ������ ���� ��: Transport-Monitoring,Transport,Transport-Updater
sub deleteService(ByRef path,ByRef tt)

	A = Split(path, " //")
	I = A(0)
	WshShell.Run I&" //SS//"&tt
	WshShell.Run I&" //DS//"&tt

	if err.number <> 0 then
		
		MsgBox "������ �������� ������ "&tt
		WScript.Quit
		
	end if	

end sub 'deleteService

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////

sub installService(pathUtmName,typeService)

	thisPath = getVbsPath()
	
	Err.Clear
	
	if typeService = "Transport-Monitoring" then
	
		WshShell.Run thisPath&"\"&pathUtmName&"\monitoring\bin\daemon.exe //IS//Transport-Monitoring --StopMode jvm --StartMode jvm --LogPrefix daemon-monitoring --Classpath "&thisPath&"\"&pathUtmName&"\monitoring\lib\attoparser-2.0.4.RELEASE.jar;"&thisPath&"\"&pathUtmName&"\monitoring\lib\bcmail-jdk15-1.45.jar;"&thisPath&"\"&pathUtmName&"\monitoring\lib\bcprov-jdk15-1.45.jar;"&thisPath&"\"&pathUtmName&"\monitoring\lib\commons-codec-1.6.jar;"&thisPath&"\"&pathUtmName&"\monitoring\lib\commons-io-2.4.jar;"&thisPath&"\"&pathUtmName&"\monitoring\lib\commons-lang-2.6.jar;"&thisPath&"\"&pathUtmName&"\monitoring\lib\commons-logging-1.2.jar;"&thisPath&"\"&pathUtmName&"\monitoring\lib\derby-10.11.1.1.jar;"&thisPath&"\"&pathUtmName&"\monitoring\lib\dom4j-2.1.0.jar;"&thisPath&"\"&pathUtmName&"\monitoring\lib\httpclient-4.5.5.jar;"&thisPath&"\"&pathUtmName&"\monitoring\lib\httpcore-4.4.9.jar;"&thisPath&"\"&pathUtmName&"\monitoring\lib\javassist-3.20.0-GA.jar;"&thisPath&"\"&pathUtmName&"\monitoring\lib\jaxen-1.1.6.jar;"&thisPath&"\"&pathUtmName&"\monitoring\lib\log4j-1.2.17.jar;"&thisPath&"\"&pathUtmName&"\monitoring\lib\ognl-3.1.12.jar;"&thisPath&"\"&pathUtmName&"\monitoring\lib\slf4j-api-1.7.2.jar;"&thisPath&"\"&pathUtmName&"\monitoring\lib\slf4j-log4j12-1.7.2.jar;"&thisPath&"\"&pathUtmName&"\monitoring\lib\terminal-pki-3.0.8.jar;"&thisPath&"\"&pathUtmName&"\monitoring\lib\terminal-util-3.0.8.jar;"&thisPath&"\"&pathUtmName&"\monitoring\lib\thymeleaf-3.0.9.RELEASE.jar;"&thisPath&"\"&pathUtmName&"\monitoring\lib\transport-monitoring-3.0.8.jar;"&thisPath&"\"&pathUtmName&"\monitoring\lib\unbescape-1.1.5.RELEASE.jar; --LogPath "&thisPath&"\"&pathUtmName&"\monitoring\l --Jvm "&thisPath&"\"&pathUtmName&"\jre\bin\client\jvm.dll --DependsOn SCardSvr --JvmOptions -Dapp.name='monitoring';-Dapp.repo="&thisPath&"\"&pathUtmName&"\monitoring\lib;-Dapp.home="&thisPath&"\"&pathUtmName&"\monitoring;-Dbasedir="&thisPath&"\"&pathUtmName&"\monitoring --StopParams 0 --StartParams "&thisPath&"\"&pathUtmName&"\monitoring\conf\transport.properties; --StartClass es.programador.transport.monitoring.Main --StdError "&thisPath&"\"&pathUtmName&"\monitoring\l\daemon-error.log --LogLevel Debug --StopMethod exit --StartMethod main --StopClass java.lang.System --StdOutput "&thisPath&"\"&pathUtmName&"\monitoring\l\daemon-output.log --Description 'Transport Terminal Monitoring' --Startup auto"
		
		if err.number <> 0 then
			
			msgBox "������ ��������� ������:"+typeService
			
		end if	
	
	end if
	
	if typeService = "Transport" then
	
		WshShell.Run thisPath&"\"&pathUtmName&"\transporter\bin\daemon.exe //IS//Transport --StopMode jvm --StartMode jvm --LogPrefix daemon-transport --Classpath "&thisPath&"\"&pathUtmName&"\transporter\lib\attoparser-2.0.4.RELEASE.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\bcmail-jdk15on-1.55.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\bcpkix-jdk15on-1.55.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\bcprov-jdk15on-1.55.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\c3p0-0.9.1.1.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\commons-codec-1.6.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\commons-configuration-1.10.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\commons-io-2.4.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\commons-lang-2.6.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\commons-logging-1.2.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\derby-10.11.1.1.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\dom4j-2.1.0.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\error_prone_annotations-2.0.2.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\guava-19.0.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\guava-probably-10a7382.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\guava-testlib-19.0.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\guava-tests-19.0.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\httpclient-4.5.5.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\httpcore-4.4.9.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\javassist-3.20.0-GA.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\javax.servlet-api-3.1.0.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\jaxen-1.1.6.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\jetty-continuation-9.3.5.v20151012.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\jetty-http-9.3.5.v20151012.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\jetty-io-9.3.5.v20151012.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\jetty-security-9.3.5.v20151012.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\jetty-server-9.3.5.v20151012.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\jetty-servlet-9.3.5.v20151012.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\jetty-servlets-9.3.5.v20151012.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\jetty-util-9.3.5.v20151012.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\jetty-webapp-9.3.5.v20151012.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\jetty-xml-9.3.5.v20151012.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\json-simple-1.1.1.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\jsr305-2.0.1.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\log4j-1.2.16.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\ognl-3.1.12.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\quartz-2.1.6.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\slf4j-api-1.7.2.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\slf4j-log4j12-1.7.2.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\terminal-backbone-3.0.8.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\terminal-conf-3.0.8.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\terminal-crypto-3.0.8.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\terminal-daemon-3.0.8.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\terminal-persist-3.0.8.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\terminal-pki-3.0.8.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\terminal-util-3.0.8.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\terminal-validator-3.0.8.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\terminal-webapp-util-3.0.8.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\terminal-ws-sender-3.0.8.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\thymeleaf-3.0.9.RELEASE.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\truth-0.28.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\unbescape-1.1.5.RELEASE.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\xercesImpl-2.11.0.jar;"&thisPath&"\"&pathUtmName&"\transporter\lib\xml-apis-1.4.01.jar; --LogPath "&thisPath&"\"&pathUtmName&"\transporter\l --Jvm "&thisPath&"\"&pathUtmName&"\jre\bin\client\jvm.dll --DependsOn SCardSvr --JvmOptions -Dderby.stream.error.file="&thisPath&"\"&pathUtmName&"\transporter\l\derby.log;-Dapp.name='transport';-Dapp.repo="&thisPath&"\"&pathUtmName&"\transporter\lib;-Dapp.home="&thisPath&"\"&pathUtmName&"\transporter;-Dbasedir="&thisPath&"\"&pathUtmName&"\transporter --StopParams 0 --StartParams "&thisPath&"\"&pathUtmName&"\transporter\conf\transport.properties --StartClass es.programador.transport.Transport --StdError "&thisPath&"\"&pathUtmName&"\transporter\l\daemon-error.log --LogLevel Info --StopMethod exit --StartMethod main --StopClass java.lang.System --StdOutput "&thisPath&"\"&pathUtmName&"\transporter\l\daemon-output.log --Description 'Transport Terminal' --Startup auto"
	
		if err.number <> 0 then
			
			msgBox "������ ��������� ������:"+typeService
			
		end if	
		
	end if
	
	if typeService = "Transport-Updater" then
	
		WshShell.Run thisPath&"\"&pathUtmName&"\updater\bin\daemon.exe //IS//Transport-Updater --StopMode jvm --StartMode jvm --LogPrefix daemon-updater --Classpath "&thisPath&"\"&pathUtmName&"\updater\lib\attoparser-2.0.4.RELEASE.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\bcmail-jdk15-1.45.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\bcprov-jdk15-1.45.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\commons-codec-1.6.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\commons-configuration-1.10.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\commons-io-2.4.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\commons-lang-2.6.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\commons-logging-1.1.1.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\derby-10.11.1.1.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\dom4j-2.1.0.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\httpclient-4.5.5.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\httpcore-4.4.9.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\javassist-3.20.0-GA.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\jaxen-1.1.6.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\log4j-1.2.17.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\ognl-3.1.12.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\slf4j-api-1.7.2.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\slf4j-log4j12-1.7.2.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\terminal-conf-3.0.8.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\terminal-daemon-3.0.8.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\terminal-pki-3.0.8.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\terminal-updater-util-3.0.8.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\terminal-util-3.0.8.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\thymeleaf-3.0.9.RELEASE.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\transport-updater-3.0.8.jar;"&thisPath&"\"&pathUtmName&"\updater\lib\unbescape-1.1.5.RELEASE.jar; --LogPath "&thisPath&"\"&pathUtmName&"\updater\l --Jvm "&thisPath&"\"&pathUtmName&"\jre\bin\client\jvm.dll --DependsOn SCardSvr --JvmOptions -Dderby.stream.error.file="&thisPath&"\"&pathUtmName&"\updater\l\derby.log;-Dapp.name='transport-updater';-Dapp.repo="&thisPath&"\"&pathUtmName&"\updater\lib;-Dapp.home="&thisPath&"\"&pathUtmName&"\updater;-Dbasedir="&thisPath&"\"&pathUtmName&"\updater --StopParams 0 --StartParams "&thisPath&"\"&pathUtmName&"\updater\conf\transport.properties;daemon-run --StartClass es.programador.transport.updater.Main --StdError "&thisPath&"\"&pathUtmName&"\updater\l\daemon-error.log --LogLevel Debug --StopMethod exit --StartMethod main --StopClass java.lang.System --StdOutput "&thisPath&"\"&pathUtmName&"\updater\l\daemon-output.log --Description 'Transport Terminal Updater' --Startup auto"
		
		if err.number <> 0 then
			
			msgBox "������ ��������� ������:"+typeService
			
		end if	
		
	end if
	
end sub 'installService

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'**������ ����� ������� ��� ��� � ���� �� �����������

function returnNamesUTM()

	dim s
	
	for each p in myArray
		
		s = s+","+p
		
	next
	
	returnNamesUTM = s

end function 'returnNamesUTM

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////



'*****��� 1 - ������ ������ , ���� ��� ����
'*****��������� �� ������� ������ �� ������� ���������������� ������ ���

Transport 				= WshShell.RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Transport\ImagePath")
TransportMonitoring 	= WshShell.RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Transport-Monitoring\ImagePath")
TransportUpdater 		= WshShell.RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Transport-Updater\ImagePath")

if err.number <> 0 then
	
	'***������ ��� � ������� , ����� ������ ���������� ������ �������
	MsgBox "������ ������ �������,�������� ������ ��� �������"
	'WScript.Quit
else
	
	'**	� ������� ��� �� ���� , ������� ������
	deleteService Transport,"Transport-Monitoring"
	deleteService TransportMonitoring,"Transport"
	deleteService TransportUpdater,"Transport-Updater"
	
end if	

'***��� 2 ��������� ������

inputName = InputBox(returnNamesUTM())
installService inputName,"Transport-Monitoring"
installService inputName,"Transport"
installService inputName,"Transport-Updater"
'WScript.Quit

MsgBox "OK"
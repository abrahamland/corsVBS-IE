
'Crear objetos WScript.Shell y FileSystemObject para escribir en registro windows y manejar archivos
Set objFSO=CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

On Error Resume Next

'Habilitar opcion CORS en internet explorer....
'Opcion 1406 ,Miscellaneous: Access data sources across domains ,0=Activar, 1=Prompt, 3=Deshabilita

''===============================Habilitar en laz zonas la opcion de configuracion 1406==============================
''Zona 1 -> Local intranet zone
objShell.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1\1406", 0, "REG_DWORD"
''Zona 2 -> Trusted sites zone
objShell.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2\1406", 0, "REG_DWORD"
''Zona 3 -> Internet zone
objShell.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\1406", 0, "REG_DWORD"
''Zona 4 -> Restricted sites Zone
objShell.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4\1406", 0, "REG_DWORD"
''===================================================================================================================

'Crear/append archivo de ejecucion de script
strNombreArchivo="corslogCDF.txt"
Set ArchivoLog = objFSO.OpenTextFile(strNombreArchivo ,8 , True)

If Err.Number <> 0 Then
	ArchivoLog.WriteLine("Error al ejecutar cors.vbs::" & Err.Description )
	''WScript.Echo "Error al ejecutar"
Else
	ArchivoLog.WriteLine("Script cors.vbs ejecutado OK, t= " & Now)
	''WScript.Echo "OK vbs"
End IF

Set objFSO = Nothing
Set objShell = Nothing






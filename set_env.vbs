'******************************************************************************
'* Fichier     : set_env.vbs                                                  *
'* Auteur      : Bruno Boissonnet                                             *
'* Version     : 1.0                                                          *
'* Description : Script qui change les variables d'environnement TEMP et TMP. *
'*               - SYSTEM TEMP = C:\Temp                                      *
'*               - SYSTEM TMP  = C:\Temp                                      *
'*               - USER TEMP   = C:\Temp                                      *
'*               - USER TEMP   = C:\Temp                                      *
'*                                                                            *
'* Remarques   :                                                              *
'*                                                                            *
'******************************************************************************


Set wshShell = CreateObject( "WScript.Shell" )
Set wshSystemEnv = wshShell.Environment( "SYSTEM" )
' Display the current value
'WScript.Echo "TEMP=" & wshSystemEnv( "TEMP" )

' Set the environment variable
wshSystemEnv( "TEMP" ) = "C:\Temp"
' Display the result
WScript.Echo "TEMP=" & wshSystemEnv( "TEMP" )

' Set the environment variable
wshSystemEnv( "TMP" ) = "C:\Temp"
' Display the result
WScript.Echo "TMP=" & wshSystemEnv( "TMP" )


Set wshSystemEnv = wshShell.Environment( "USER" )

' Display the current value
'WScript.Echo "TEMP=" & wshSystemEnv( "TEMP" )

' Set the environment variable
wshSystemEnv( "TEMP" ) = "C:\Temp"
' Display the result
WScript.Echo "TEMP=" & wshSystemEnv( "TEMP" )

' Set the environment variable
wshSystemEnv( "TMP" ) = "C:\Temp"
' Display the result
WScript.Echo "TMP=" & wshSystemEnv( "TMP" )


Set wshSystemEnv = Nothing
Set wshShell     = Nothing
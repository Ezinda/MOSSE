Attribute VB_Name = "Module1"
Option Explicit

'constantes para obtener la informaci�n regional
Public Const LOCALE_FONTSIGNATURE = &H58
Public Const LOCALE_ICENTURY = &H24
Public Const LOCALE_ICOUNTRY = &H5 'c�digo de pa�s
Public Const LOCALE_ICURRDIGITS = &H19 'n� de decimales en las monedas
Public Const LOCALE_ICURRENCY = &H1B 'posici�n del simbolo de moneda respecto al n�mero,0=delante, 1=detr�s, 2=delante con un blanco, 3=detras con un blanco
Public Const LOCALE_IDATE = &H21
Public Const LOCALE_IDAYLZERO = &H26 '1=d�as con dos d�gitos en fecha corta
Public Const LOCALE_IDEFAULTCODEPAGE = &HB 'p�gina de c�digos por defecto
Public Const LOCALE_IDEFAULTCOUNTRY = &HA 'c�digo de pa�s por defecto
Public Const LOCALE_IDEFAULTLANGUAGE = &H9 'codigo de lenguaje por defecto
Public Const LOCALE_IDIGITS = &H11 'n� de decimales en los numeros
Public Const LOCALE_IINTLCURRDIGITS = &H1A
Public Const LOCALE_ILANGUAGE = &H1 'codigo del lenguaje
Public Const LOCALE_ILDATE = &H22
Public Const LOCALE_ILZERO = &H12
Public Const LOCALE_IMEASURE = &HD 'sistema de medida, 0=metrico, 1 =EE.UU.
Public Const LOCALE_IMONLZERO = &H27
Public Const LOCALE_INEGCURR = &H1C 'formato n� negativo en las monedas
Public Const LOCALE_INEGSEPBYSPACE = &H57 'un espacio entre el n� y la moneda en los negativos
Public Const LOCALE_INEGSIGNPOSN = &H53 'posicion del signo en las monedas negativas, 0=no se pone, 1=antes del numero, 2=despues del numero,3=antes de la moneda,4=despues de la monea
Public Const LOCALE_INEGSYMPRECEDES = &H56
Public Const LOCALE_IPOSSEPBYSPACE = &H55
Public Const LOCALE_IPOSSIGNPOSN = &H52
Public Const LOCALE_IPOSSYMPRECEDES = &H54
Public Const LOCALE_ITIME = &H23
Public Const LOCALE_ITLZERO = &H25 '1=horas con dos digitos
Public Const LOCALE_NOUSEROVERRIDE = &H80000000
Public Const LOCALE_S1159 = &H28 'simbolo a.m.
Public Const LOCALE_S2359 = &H29 'simbolo p.m.
Public Const LOCALE_SABBREVCTRYNAME = &H7 'nombre abreviado del pa�s
Public Const LOCALE_SABBREVDAYNAME1 = &H31 'nombre abreviado de los d�as de la semana
Public Const LOCALE_SABBREVDAYNAME2 = &H32 'en el idioma del pa�s
Public Const LOCALE_SABBREVDAYNAME3 = &H33
Public Const LOCALE_SABBREVDAYNAME4 = &H34
Public Const LOCALE_SABBREVDAYNAME5 = &H35
Public Const LOCALE_SABBREVDAYNAME6 = &H36
Public Const LOCALE_SABBREVDAYNAME7 = &H37
Public Const LOCALE_SABBREVLANGNAME = &H3 'nombre a breviado del lenguaje
Public Const LOCALE_SABBREVMONTHNAME1 = &H44  'nombre abreviado de los meses del a�o
Public Const LOCALE_SABBREVMONTHNAME10 = &H4D 'en el idioma del pa�s
Public Const LOCALE_SABBREVMONTHNAME11 = &H4E
Public Const LOCALE_SABBREVMONTHNAME12 = &H4F
Public Const LOCALE_SABBREVMONTHNAME13 = &H100F
Public Const LOCALE_SABBREVMONTHNAME2 = &H45
Public Const LOCALE_SABBREVMONTHNAME3 = &H46
Public Const LOCALE_SABBREVMONTHNAME4 = &H47
Public Const LOCALE_SABBREVMONTHNAME5 = &H48
Public Const LOCALE_SABBREVMONTHNAME6 = &H49
Public Const LOCALE_SABBREVMONTHNAME7 = &H4A
Public Const LOCALE_SABBREVMONTHNAME8 = &H4B
Public Const LOCALE_SABBREVMONTHNAME9 = &H4C
Public Const LOCALE_SCOUNTRY = &H6 'nombre del pa�s en ingl�s
Public Const LOCALE_SCURRENCY = &H14 's�mbolo de la moneda
Public Const LOCALE_SDATE = &H1D 'separador de fechas
Public Const LOCALE_SDAYNAME1 = &H2A 'nombre de los d�as d�a de la semana
Public Const LOCALE_SDAYNAME2 = &H2B 'en el idioma del pa�s
Public Const LOCALE_SDAYNAME3 = &H2C
Public Const LOCALE_SDAYNAME4 = &H2D
Public Const LOCALE_SDAYNAME5 = &H2E
Public Const LOCALE_SDAYNAME6 = &H2F
Public Const LOCALE_SDAYNAME7 = &H30
Public Const LOCALE_SDECIMAL = &HE 'separador decimal
Public Const LOCALE_SENGCOUNTRY = &H1002
Public Const LOCALE_SENGLANGUAGE = &H1001
Public Const LOCALE_SGROUPING = &H10 'n� de d�gitos en grupo
Public Const LOCALE_SINTLSYMBOL = &H15 'simbolo internacional del pais
Public Const LOCALE_SLANGUAGE = &H2 'lenguaje selecionado en conf.reg.
Public Const LOCALE_SLIST = &HC 'separador de listas
Public Const LOCALE_SLONGDATE = &H20 'formato de fecha larga
Public Const LOCALE_SMONDECIMALSEP = &H16 'separador decimal en las monedas
Public Const LOCALE_SMONGROUPING = &H18 'n� de d�gitos en grupo para las monedas
Public Const LOCALE_SMONTHNAME1 = &H38  'nombres de los meses
Public Const LOCALE_SMONTHNAME10 = &H41 'en el idioma del pa�s
Public Const LOCALE_SMONTHNAME11 = &H42
Public Const LOCALE_SMONTHNAME12 = &H43
Public Const LOCALE_SMONTHNAME2 = &H39
Public Const LOCALE_SMONTHNAME3 = &H3A
Public Const LOCALE_SMONTHNAME4 = &H3B
Public Const LOCALE_SMONTHNAME5 = &H3C
Public Const LOCALE_SMONTHNAME6 = &H3D
Public Const LOCALE_SMONTHNAME7 = &H3E
Public Const LOCALE_SMONTHNAME8 = &H3F
Public Const LOCALE_SMONTHNAME9 = &H40
Public Const LOCALE_SMONTHOUSANDSEP = &H17 'separador de miles en las monedas
Public Const LOCALE_SNATIVECTRYNAME = &H8 'nombre del pa�s en el idioma del pa�s
Public Const LOCALE_SNATIVEDIGITS = &H13 'digitos empleados en el pa�s
Public Const LOCALE_SNATIVELANGNAME = &H4 'idioma del pa�s en el idioma del pa�s
Public Const LOCALE_SNEGATIVESIGN = &H51 'simbolo de signo negativo
Public Const LOCALE_SPOSITIVESIGN = &H50 'simbolo de signo positivo
Public Const LOCALE_SSHORTDATE = &H1F 'formato de fecha corta
Public Const LOCALE_STHOUSAND = &HF 'separador de miles
Public Const LOCALE_STIME = &H1E 'separador de horas
Public Const LOCALE_STIMEFORMAT = &H1003 'formato de horas
Public Const LOCALE_SYSTEM_DEFAULT = &H800 'presentar informaci�n del sistema
Public Const LOCALE_USER_DEFAULT = &H400 'presentar informaci�n del usuario

's�lo se pueden modificar los siguientes valores :
'LOCALE_SDATE
'LOCALE_ICURRDIGITS
'LOCALE_SDECIMAL
'LOCALE_ICURRENCY
'LOCALE_SGROUPING
'LOCALE_IDIGITS
'LOCALE_SLIST
'LOCALE_SLONGDATE
'LOCALE_SMONDECIMALSEP
'LOCALE_ILZERO
'LOCALE_SMONGROUPING
'LOCALE_IMEASURE
'LOCALE_SMONTHOUSANDSEP
'LOCALE_INEGCURR
'LOCALE_SNEGATIVESIGN
'LOCALE_INEGNUMBER
'LOCALE_SPOSITIVESIGN
'LOCALE_ITIME
'LOCALE_SSHORTDATE
'LOCALE_S1159
'LOCALE_STHOUSAND
'LOCALE_S2359
'LOCALE_STIME
'LOCALE_SCURRENCY
'LOCALE_STIMEFORMAT


Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long


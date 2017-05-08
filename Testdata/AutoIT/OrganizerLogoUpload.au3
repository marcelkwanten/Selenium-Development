#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>
#include <Constants.au3>

Opt("WinTitleMatchMode", 2)

$sFileUpload = "D:\Users\mkwanten\Documents\Opdrachten\Ignation\SeleniumFrameworkTemplate2\Testdata\AutoIT\FileUpload.exe"

Send("!{TAB}")

	 If WinWaitActive("Open", "", 1500) Then
			Sleep(500)
			Send("!d")
			Sleep(500)
			Send("D:\Users\mkwanten\Documents\Opdrachten\Ignation\Testafbeeldingen")
			Send("{ENTER}")
			Send("{F6 8}")
			Sleep(500)
			Send("ignation_logo_big.png")
			Send("{ENTER}")

	 EndIf

	 Sleep(5000)

	 Run($sFileUpload)
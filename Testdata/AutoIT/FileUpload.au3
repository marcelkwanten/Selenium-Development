#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>
#include <Constants.au3>

Opt("WinTitleMatchMode", 2)

Send("!{TAB}")

	If WinWaitActive("Open", "", 1500) Then
			Sleep(500)
			Send("!d")
			Sleep(500)
			Send("D:\Users\mkwanten\Documents\Opdrachten\Ignation\Testafbeeldingen")
			Send("{ENTER}")
			Send("{F6 8}")
			Sleep(500)
			Send("800x800.PNG")
			Send("{ENTER}")

	EndIf
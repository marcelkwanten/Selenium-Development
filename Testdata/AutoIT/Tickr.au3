#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>
#include <Constants.au3>
#include <Excel.au3>
#include <MsgBoxConstants.au3>

#AutoIt3Wrapper_icon=favicon.ico
Opt("WinTitleMatchMode", 2)


;PAS HIER DE BESTANDSLOCATIES AAN
		$sEmailLocatie 			= "D:\Users\mkwanten\Documents\Opdrachten\Ignation\SeleniumFrameworkTemplate2\Testdata\Testcases\import\email"
		$sStarter 				= "D:\Users\mkwanten\Documents\Opdrachten\Ignation\SeleniumFrameworkTemplate2"
		$sTestSuite 			= "D:\Users\mkwanten\Documents\Opdrachten\Ignation\SeleniumFrameworkTemplate2\Testdata\Testsuites\testsuite.xls"
		$sOrganizerLogoUpload 	= "D:\Users\mkwanten\Documents\Opdrachten\Ignation\SeleniumFrameworkTemplate2\Testdata\AutoIT\OrganizerLogoUpload.exe"
		$sDevelopment 			= "http://95.85.29.126:665/"
		$sAcceptatie 			= "http://146.185.161.50:665/"


; 	GUI
		Local $GUI = GUICreate("Tickr Test Automation", 300, 300)
		$sTitelKolom1 = GUICtrlCreateLabel("Vul hier de cijfers in die achter de e-mailadressen komen:", 15, 10, 280)

;LABEL GETAL 1
		$sLabel_Getal 	= GUICtrlCreateLabel("Getal 1: ", 15, 35, 280, 20)
;INPUT GETAL 1
		$iGetal 		= GUICtrlCreateInput("", 15, 50, 50, 22)

;LABEL GETAL 2
		$sLabel_Getal2 	= GUICtrlCreateLabel("Getal 2: ", 100, 35, 280, 20)
;INPUT GETAL 2
		$iGetal2 		= GUICtrlCreateInput("", 105, 50, 50, 22)

;WIJZIG E-MAILADRESSEN BUTTON
		$RUN_1 			= GUICtrlCreateButton("Wijzig E-mailadressen", 15, 75, 125, 30)

;LABEL URL
		$sLabel_URL 	= GUICtrlCreateLabel("URL Website: ", 15, 120, 280, 20)
;INPUT URL
		$sInput_URL 	= GUICtrlCreateCombo("Kies URL", 15, 140, 270, 20)
		GUICtrlSetData(-1, "http://95.85.29.126:665|http://146.185.161.50:665/|https://tickr.io/", "")

;WIJZIG URL BUTTON
		$RUN_URL 		= GUICtrlCreateButton("Wijzig URL", 15, 165, 125, 30)


;OPEN TESTSUITE
		$RUN_TS 		= GUICtrlCreateButton("Open Testsuite.xls", 15, 225, 125, 30)

;TESTSCRIPT UITVOEREN BUTTON
		$RUN_2 			= GUICtrlCreateButton("Start Test", 15, 260, 125, 30)

;CLOSE
		$CLOSE 			= GUICtrlCreateButton("Sluit", 160, 260, 125, 30)



;GUI MESSAGE LOOP
	GUISetState()

		local $sComboRead = ""




	While 1
		$MSG = GUIGetMsg()
		Select
			Case $MSG = $GUI_EVENT_CLOSE
			Exit



;E-MAILADRESSEN WIJZIGEN

			Case $MSG = $RUN_1
				$iNumber = GUICtrlRead($iGetal)
				$iNumber2 = GUICtrlRead($iGetal2)



				_Excel_BookOpen(_Excel_Open(), $sEmailLocatie & "\create\consumer.xls")

					If WinWaitActive("consumer", "", 10) then
						Local $oExcel = ObjGet("", "Excel.Application")
						$oExcel.Range("D2").activate
						Send("testignation{+}consumer" & $iNumber & "@gmail.com")
						Send("^s")
						Sleep(500)
						Send("!{f4}")
					EndIf



				_Excel_BookOpen(_Excel_Open(), $sEmailLocatie & "\forgotpassword\consumer.xls")

					If WinWaitActive("consumer", "", 10) then
						Local $oExcel = ObjGet("", "Excel.Application")
						$oExcel.Range("D2").activate
						Send("testignation{+}consumer" & $iNumber & "@gmail.com")
						Send("^s")
						Sleep(500)
						Send("!{f4}")
					EndIf


				_Excel_BookOpen(_Excel_Open(), $sEmailLocatie & "\iframe\consumer.xls")

					If WinWaitActive("consumer", "", 10) then
						Local $oExcel = ObjGet("", "Excel.Application")
						$oExcel.Range("D2").activate
						Send("testignation{+}consumer" & $iNumber & "@gmail.com")
						Send("^s")
						Sleep(500)
						Send("!{f4}")
					EndIf



				_Excel_BookOpen(_Excel_Open(), $sEmailLocatie & "\login\consumer.xls")

					If WinWaitActive("consumer", "", 10) then
						Local $oExcel = ObjGet("", "Excel.Application")
						$oExcel.Range("D2").activate
						Send("testignation{+}consumer" & $iNumber & "@gmail.com")
						Send("^s")
						Sleep(500)
						Send("!{f4}")
					EndIf



				_Excel_BookOpen(_Excel_Open(), $sEmailLocatie & "\create\consumer2.xls")

					If WinWaitActive("consumer2", "", 10) then
						Local $oExcel = ObjGet("", "Excel.Application")
						$oExcel.Range("D2").activate
						Send("testignation{+}consumer" & $iNumber2 & "@gmail.com")
						Send("^s")
						Sleep(500)
						Send("!{f4}")
					EndIf



				_Excel_BookOpen(_Excel_Open(), $sEmailLocatie & "\login\consumer2.xls")

					If WinWaitActive("consumer2", "", 10) then
						Local $oExcel = ObjGet("", "Excel.Application")
						$oExcel.Range("D2").activate
						Send("testignation{+}consumer" & $iNumber2 & "@gmail.com")
						Send("^s")
						Sleep(500)
						Send("!{f4}")
					EndIf



				_Excel_BookOpen(_Excel_Open(), $sEmailLocatie & "\mytickets\consumer2.xls")

					If WinWaitActive("consumer2", "", 10) then
						Local $oExcel = ObjGet("", "Excel.Application")
						$oExcel.Range("D2").activate
						Send("testignation{+}consumer" & $iNumber2 & "@gmail.com")
						Send("^s")
						Sleep(500)
						Send("!{f4}")
					EndIf


				_Excel_BookOpen(_Excel_Open(), $sEmailLocatie & "\create\inactiveconsumer.xls")

					If WinWaitActive("inactiveconsumer", "", 10) then
						Local $oExcel = ObjGet("", "Excel.Application")
						$oExcel.Range("D2").activate
						Send("testignation{+}inactiveconsumer" & $iNumber & "@gmail.com")
						Send("^s")
						Sleep(500)
						Send("!{f4}")
					EndIf



				_Excel_BookOpen(_Excel_Open(), $sEmailLocatie & "\login\inactiveconsumer.xls")

					If WinWaitActive("inactiveconsumer", "", 10) then
						Local $oExcel = ObjGet("", "Excel.Application")
						$oExcel.Range("D2").activate
						Send("testignation{+}inactiveconsumer" & $iNumber & "@gmail.com")
						Send("^s")
						Sleep(500)
						Send("!{f4}")
					EndIf



				_Excel_BookOpen(_Excel_Open(), $sEmailLocatie & "\create\organizer.xls")

					If WinWaitActive("organizer", "", 10) then
						Local $oExcel = ObjGet("", "Excel.Application")
						$oExcel.Range("D2").activate
						Send("testignation{+}organizer" & $iNumber & "@gmail.com")
						Send("^s")
						Sleep(500)
						Send("!{f4}")
					EndIf


				_Excel_BookOpen(_Excel_Open(), $sEmailLocatie & "\forgotpassword\organizer.xls")

					If WinWaitActive("organizer", "", 10) then
						Local $oExcel = ObjGet("", "Excel.Application")
						$oExcel.Range("D2").activate
						Send("testignation{+}organizer" & $iNumber & "@gmail.com")
						Send("^s")
						Sleep(500)
						Send("!{f4}")
					EndIf


				_Excel_BookOpen(_Excel_Open(), $sEmailLocatie & "\login\organizer.xls")

					If WinWaitActive("organizer", "", 10) then
						Local $oExcel = ObjGet("", "Excel.Application")
						$oExcel.Range("D2").activate
						Send("testignation{+}organizer" & $iNumber & "@gmail.com")
						Send("^s")
						Sleep(500)
						Send("!{f4}")
					EndIf



				_Excel_BookOpen(_Excel_Open(), $sEmailLocatie & "\create\inactiveorganizer.xls")

					If WinWaitActive("inactiveorganizer", "", 10) then
						Local $oExcel = ObjGet("", "Excel.Application")
						$oExcel.Range("D2").activate
						Send("testignation{+}inactiveorganizer" & $iNumber & "@gmail.com")
						Send("^s")
						Sleep(500)
						Send("!{f4}")
					EndIf



				_Excel_BookOpen(_Excel_Open(), $sEmailLocatie & "\login\inactiveorganizer.xls")

					If WinWaitActive("inactiveorganizer", "", 10) then
						Local $oExcel = ObjGet("", "Excel.Application")
						$oExcel.Range("D2").activate
						Send("testignation{+}inactiveorganizer" & $iNumber & "@gmail.com")
						Send("^s")
						Sleep(500)
						Send("!{f4}")
					EndIf




				_Excel_BookOpen(_Excel_Open(), $sEmailLocatie & "\iframe\guest.xls")

					If WinWaitActive("guest", "", 10) then
						Local $oExcel = ObjGet("", "Excel.Application")
						$oExcel.Range("D2").activate
						Send("testignation{+}guest" & $iNumber & "@gmail.com")
						Send("^s")
						Sleep(500)
						Send("!{f4}")
					EndIf



				_Excel_BookOpen(_Excel_Open(), $sEmailLocatie & "\iframe\newuser.xls")

					If WinWaitActive("newuser", "", 10) then
						Local $oExcel = ObjGet("", "Excel.Application")
						$oExcel.Range("D2").activate
						Send("testignation{+}newuser" & $iNumber & "@gmail.com")
						Send("^s")
						Sleep(500)
						Send("!{f4}")
					EndIf



				_Excel_BookOpen(_Excel_Open(), $sEmailLocatie & "\create\doorman.xls")

					If WinWaitActive("doorman", "", 10) then
						Local $oExcel = ObjGet("", "Excel.Application")
						$oExcel.Range("D2").activate
						Send("testignation{+}doorman" & $iNumber & "@gmail.com")
						Send("^s")
						Sleep(500)
						Send("!{f4}")
					EndIf



				_Excel_BookOpen(_Excel_Open(), $sEmailLocatie & "\create\eventmanager.xls")

					If WinWaitActive("eventmanager", "", 10) then
						Local $oExcel = ObjGet("", "Excel.Application")
						$oExcel.Range("D2").activate
						Send("testignation{+}manager" & $iNumber & "@gmail.com")
						Send("^s")
						Sleep(500)
						Send("!{f4}")
					EndIf

;URL'S INVOEREN
			Case $MSG = $RUN_URL

			$URL = GUICtrlRead($sInput_URL)
				_Excel_BookOpen(_Excel_Open(), $sTestSuite)
					Sleep(1000)
					If WinWaitActive("testsuite", "", 10) then
						Local $oExcel = ObjGet("", "Excel.Application")
						$oExcel.Range("B3").activate

						Local $i = 0;
							While $i <= 29
                           Send($URL)
                           Send("{ENTER}")
                           $i = $i + 1
							WEnd

						Send("^s")
						Sleep(500)
						Send("!{f4}")
					EndIf


;TESTSUITE OPENEN
			Case $MSG = $RUN_TS
				_Excel_BookOpen(_Excel_Open(), $sTestSuite)
					If WinWaitActive("testsuite", "", 10) then
						Local $oExcel = ObjGet("", "Excel.Application")
						$oExcel.Range("D2").activate
					EndIf



;TEST STARTEN

			Case $MSG = $RUN_2

					If ProcessExists("OrganizerLogoUpload.exe") Then
						MsgBox("", "Image Upload", "Image upload is already running")
					Else
						Run($sOrganizerLogoUpload)
						Sleep(3000)
					EndIf

						Run("C:\Windows\System32\cmd.exe")
							If WinWaitActive("Administrator", "",10) Then
								Send("pushd " & $sStarter)
								Send("{enter}")
								Send("java -jar Starter.jar suite " & "testsuite.xls")
								Send("{enter}")
							EndIf
				Sleep(3000)
				While ProcessExists("cmd.exe")
					If Not ProcessExists("cmd.exe") Then
						ProcessClose("OrganizerLogoUpload.exe")
						ProcessClose("FileUpload.exe")

					EndIf
				WEnd

;APP AFSLUITEN

			Case $MSG = $CLOSE
				Send("!{F4}")

		EndSelect
	WEnd




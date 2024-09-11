#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=
#AutoIt3Wrapper_Res_Comment=
#AutoIt3Wrapper_Res_Description=Rebuilds Outlook Autocomplete Address List using unique addresses from email messages in Sent Items folder.
#AutoIt3Wrapper_Res_Fileversion=07.09.2020.1800
#AutoIt3Wrapper_Res_ProductName=Rebuild Outlook Autocomplete Address List
#AutoIt3Wrapper_Res_ProductVersion=07.09.2020.1800
#AutoIt3Wrapper_Res_LegalCopyright=© 2020 Nathan Larsen
#AutoIt3Wrapper_Res_Language=1033
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         Nathan Larsen
 Finalized:      07.09.2020 at 1800

 Script Function:
	Rebuild Outlook Autocomplete Address List using unique addresses from email messages in Sent Items folder

#ce ----------------------------------------------------------------------------

AutoItSetOption("MustDeclareVars", 1)

#include <Date.au3>

Global $objOutlook = ObjCreate("Outlook.Application")
Global $objNamespace = $objOutlook.GetNamespace("MAPI")
Global $arrAddresses[1]
Const $olMailItem = 0

; Create address list of unique addresses from emails in Sent Items folder
Global $objSentItemsFolder = $objNamespace.GetDefaultFolder(5)

; Create progress bar to track progress of email processing
ProgressOn("Address List Creation Status", "", "Working...")

; Variable used for progress bar incrementation
Global $intMsgCount = 0

Global $objEmail
For $objEmail In $objSentItemsFolder.Items

	; Increment progress bar variable
	$intMsgCount = $intMsgCount + 1

	; Calculate progess bar percentage
	ProgressSet(($intMsgCount*100)/$objSentItemsFolder.Items.Count)

	Global $objRecipient
	For $objRecipient In $objEmail.Recipients

		Global $blnAdd = True

		Global $intNum
		For $intNum = 0 to UBound($arrAddresses) - 1

			If $arrAddresses[$intNum] = $objRecipient.address Or StringInStr($objRecipient.address, "@") = 0 Then

				$blnAdd = False

			EndIf

		Next

		If $blnAdd = True Then

			ReDim $arrAddresses[UBound($arrAddresses) + 1]
			$arrAddresses[UBound($arrAddresses) - 2] = $objRecipient.address

		EndIf

	Next

Next

; Display progress bar with done status and turn off progress bar
ProgressSet(100, "Done!")
Sleep(750)
ProgressOff()

; Remove blank entry in array
ReDim $arrAddresses[UBound($arrAddresses) - 1]

; Create email message
Global $olMessage = $objOutlook.CreateItem($olMailItem)
$olMessage.Subject = "Test Email"
$olMessage.Body = "This is a test email message."

; Add unique addresses to email message that were retrieved from messages in Sent Items folder
Global $intAddressNum
For $intAddressNum = 0 to UBound($arrAddresses) - 1

	$olMessage.RecipIents.Add($arrAddresses[$intAddressNum])

Next

; Send email message with twelve hour delay
$olMessage.DeferredDeliveryTime = _DateTimeFormat(_DateAdd('h', 12, _NowCalc()), 0)
$olMessage.Send

; Open first message in Outlook Outbox
Global $objOutboxFolder = $objNamespace.GetDefaultFolder(4)
Global $objFirstOutboxMessage = $objOutboxFolder.Items(1)
$objFirstOutboxMessage.Display

; Send message so addresses will post to Autocomplete List
WinWaitActive("Test Email")
Send("!s")

; Delete first message in Outlook Outbox folder
$objOutboxFolder = $objNamespace.GetDefaultFolder(4)
$objFirstOutboxMessage = $objOutboxFolder.Items(1)
$objFirstOutboxMessage.Delete

; Delete first message in Outlook Deleted Items folder
Global $objDeletedItemsFolder = $objNamespace.GetDefaultFolder(3)
Global $objFirstDeletedItemsMessage = $objDeletedItemsFolder.Items(1)
$objFirstDeletedItemsMessage.Delete

; Popup message to let user know that script completed
MsgBox(64, "Completed Rebuild", "Completed rebuilding Outlook Autocomplete Address List.")
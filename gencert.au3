;#include <IE.au3>
#include <GuiListView.au3>

Const $LoginPage = "Entrust Authority (TM) Security Manager Administration Log in"
Const $MainPage = "Security Manager Administration - (RBS FM Administrator) cn="; only need to match from the start
;$MainPage = "Security Manager Administration - (RBS FM Administrator) cn=Christopher Yeo, ou=RBS FM, ou=External Test, o=The Royal Bank of Scotland Group, c=gb"
Const $CertifyPage = "New Certification - Windows Internet Explorer"

$ssoIdsText = "testchris2.apiuat,testchris3.apiuat"
$genPrivateKeys = True
$passwd = "S@m20134"


If $CmdLine[0] = 0 Then
   ConsoleWrite("Running in test mode." & @CRLF)   
ElseIf $CmdLine[0] <> 3 Then
   ConsoleWrite("Please use <comma delimited ssoids i.e testchris2.apiuat,testchris3.apiuat> <T/F to generate private keys> <entrust login>." & @CRLF)   
   Exit
ElseIf
   $ssoIdsText = $CmdLine[1]
   If  $CmdLine[2] = "F" Then
	  $genPrivateKeys = False
   EndIf
   $passwd = $CmdLine[3]
EndIf

$text = "Using the following parameters:-" & @CRLF & _
   "	ssoIdsText:-->" & $ssoIdsText & @CRLF & _
   "	genPrivateKeys:-->" & $genPrivateKeys & @CRLF &  _
   "	passwd:-->" & $passwd & @CRLF
$reply = MsgBox (1, "Ok to proceed?", $text)
If $reply = 2 Then
   Exit
EndIf

   

$ssoIds = StringSplit( $ssoIdsText, "," )
$max = UBound($ssoIds)
Global $references[$max]
Global $passcodes[$max]

ConsoleWrite("Launch App." & @CRLF)
;Run("C:\app\Entrust\SMA UAT\entrustra.exe")
ShellExecute ("entrustra.exe")
WinWaitActive($LoginPage)
ControlFocus ($LoginPage, "", 1064)
Send($passwd & "{ENTER}")
Sleep(15000)
WinWaitActive($MainPage,"",10)   

ConsoleWrite("App login." & @CRLF)

For $i = 1 to $max -1
   
   ConsoleWrite("Trying user " & $ssoIds[$i] & @CRLF)      
	  
   Send("!u{DOWN}{RIGHT}{ENTER}")
   Send($ssoIds[$i] & "{ENTER}")
   
   ;ConsoleWrite("Timeout wait for ops.")
   ;WinWaitActive("Operation Completed Successfully","",10)
   $ret = WinWaitActive("Operation Completed Successfully","",20)
   ConsoleWrite($ret & @CRLF)
   If $ret <> 0 Then
	  ConsoleWrite("Found user." & @CRLF)
	  Send("{ENTER}")
	  ControlFocus ($MainPage, "", 59649)
 	  Local $list = ControlGetHandle ( $MainPage, "", 59649 )
	  ConsoleWrite("Found List ->" & $list & @CRLF)
	  _GUICtrlListView_ClickItem($list, 0)

	  ;Send("{DOWN}")
	  ;Sleep(2000)
	  ;Send("+{F10}")
	  ;Send("{ESC}")
	  Send("!u{DOWN}{DOWN}{DOWN}{RIGHT}")
	  Send("{DOWN}{DOWN}{DOWN}{DOWN}{ENTER}") ;Reissue Activation Code
	  $ret = WinWaitActive("Operation Completed Successfully","",5)
	  If $ret <> 0 Then
		 ConsoleWrite("Reissue OK.")
	  Else
		 ConsoleWrite("Trying Recovery.")
		 Send("!u{DOWN}{DOWN}{DOWN}{RIGHT}")
		 Send("{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{ENTER}") ;Begin Key Recovery
		 WinWaitActive("Authorization Required")
		 Send("{DOWN}")
		 Send($passwd & "{ENTER}")	  
		 ConsoleWrite("Recovery OK." & @CRLF)
	  EndIf
   Else
	  ConsoleWrite("Timeout wait for ops." & @CRLF)
	  WinWaitActive("Operation failed")
	  Send("{ENTER}")
	  Send("!u{ENTER}")
	  WinWaitActive("New User")
	  ControlFocus ("New User", "", 1476)
	  Send("{DOWN}")
	  Send($ssoIds[$i] & "{TAB}")
	  Send($ssoIds[$i] & "{ENTER}")
	  ConsoleWrite("Gen New.")
   EndIf
   
   WinWaitActive("Operation Completed Successfully")
   $codes = ControlGetText("Operation Completed Successfully","",1088)
   
   ConsoleWrite("Got codes: -->" & $codes & @CRLF)
   
   Send("{ENTER}")

;;;;;;;;;;;;;;;;;Sample;;;;;;;;;;;;;;;;;;;
;$codes = 'User has been added successfully: cn=testchris.apiuat, ou=RBS FM, ou=External Test, o=The Royal Bank of Scotland Group, c=gb'  & @CRLF & _
;'Distribute these activation codes securely to the user.'  & _
;'Reference number: 63451649'  & _
;'Authorization code: 993W-EI98-ZTOL'  & _
;'The activation codes can also be retrieved from the Activation Codes page of the user''s properties.'

   $refcount = StringInStr($codes,"Reference number: ")
   $references[$i] = StringMid($codes,  $refcount + 18,8)
   ;$codes[$i][1] = $ref
   $codecount = StringInStr($codes,"Authorization code: ")
   $passcodes[$i] = StringMid($codes,  $codecount + 20,14)
   ;$codes[$i][2] = $code
   ConsoleWrite("User codes:-" & $ssoIds[$i] & ' ' & $references[$i] & ' ' & $passcodes[$i])

Next

Send("!f{DOWN}{DOWN}{DOWN}{DOWN}{ENTER}")
ConsoleWrite("Shutdown entrust" & @CRLF)

IF $genPrivateKeys = True  Then

   For $j = 1 to $max -1
	  ConsoleWrite("SSO username:-" & $ssoIds[$j] & @CRLF & _
		 " Reference number:" & $references[$j] & @CRLF & _ 
		 " Authorization code" & $passcodes[$j] & @CRLF)
	  ;Run("C:\Program Files\Internet Explorer\iexplore.exe https://uat.rbsm.com/Logon/Certify.aspx")
	  ShellExecute("iexplore.exe","https://uat.rbsm.com/Logon/Certify.aspx") 
	  WinWaitActive($CertifyPage)
	  $pid = WinGetProcess($CertifyPage)
	  ConsoleWrite("PID is " & $pid & @CRLF)
	  Sleep(5000)	  
	  Send("{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}")
	  Send($ssoIds[$j] & "{TAB}Password1{TAB}Password1{TAB}")
	  Send($references[$j] & "{TAB}")
	  Send($passcodes[$j] & "{ENTER}")
	  
	  $ret = WinWaitActive("Security Warning","",10) 
	  If $ret <> 0 Then
		 Send("{TAB}{TAB}{ENTER}")
	  EndIf

	 ConsoleWrite("Generated private key ok."  & @CRLF)
	 Sleep(20000)
	 
	 ;ProcessClose($pid)
	 ;ProcessWaitClose($pid)
	 ;ConsoleWrite("PID Close."  & @CRLF)
	 
  Next
EndIf



;;;; Open IE9 and google search for "Mike Pennington"
;_IEErrorHandlerRegister("MyErrFunc")
;_IEErrorHandlerRegister()
;Local $oIE = _IECreate("https://uat.rbsm.com/Logon/Certify.aspx")
;; Find the form named 'f'
;WinWaitActive("New Certification - Windows Internet Explorer")
;Sleep(5000)
;Local $oForm = _IEFormGetObjByName($oIE, "frmCertify")
;; Find the form field in the form
;Local $oQuery = _IEFormElementGetObjByName($oForm, "certifyView$txtUsername")
;; Fill the form field assigned to $oQuery with the search text
;_IEFormElementSetValue($oQuery, "Test")
;_IEFormSubmit($oForm)
		 



;Func MyErrFunc()
    ; Important: the error object variable MUST be named $oIEErrorHandler
 ;   Local $ErrorScriptline = $oIEErrorHandler.scriptline
  ;  Local $ErrorNumber = $oIEErrorHandler.number
   ; Local $ErrorNumberHex = Hex($oIEErrorHandler.number, 8)
    ;Local $ErrorDescription = StringStripWS($oIEErrorHandler.description, 2)
;    Local $ErrorWinDescription = StringStripWS($oIEErrorHandler.WinDescription, 2)
 ;   Local $ErrorSource = $oIEErrorHandler.Source
  ;  Local $ErrorHelpFile = $oIEErrorHandler.HelpFile
   ; Local $ErrorHelpContext = $oIEErrorHandler.HelpContext
    ;Local $ErrorLastDllError = $oIEErrorHandler.LastDllError
;    Local $ErrorOutput = ""
 ;   $ErrorOutput &= "--> COM Error Encountered in " & @ScriptName & @CR
  ;  $ErrorOutput &= "----> $ErrorScriptline = " & $ErrorScriptline & @CR
   ; $ErrorOutput &= "----> $ErrorNumberHex = " & $ErrorNumberHex & @CR
    ;$ErrorOutput &= "----> $ErrorNumber = " & $ErrorNumber & @CR
;    $ErrorOutput &= "----> $ErrorWinDescription = " & $ErrorWinDescription & @CR
 ;   $ErrorOutput &= "----> $ErrorDescription = " & $ErrorDescription & @CR
  ;  $ErrorOutput &= "----> $ErrorSource = " & $ErrorSource & @CR
   ; $ErrorOutput &= "----> $ErrorHelpFile = " & $ErrorHelpFile & @CR
;    $ErrorOutput &= "----> $ErrorHelpContext = " & $ErrorHelpContext & @CR
 ;   $ErrorOutput &= "----> $ErrorLastDllError = " & $ErrorLastDllError
  ;  MsgBox(0, "COM Error", $ErrorOutput)
   ; SetError(1)
    ;Return
;EndFunc   ;==>MyErrFunc




#RequireAdmin
#include "Network.au3"
#include <File.au3>
#Include <array.au3>
#include <String.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <GUIListBox.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#Region ### START Koda GUI section ### Form=C:\Users\CARNET-PERSO\Desktop\NCL\Form1.kxf
$main = GUICreate("42", 186, 314, 334, 181)
$next = GUICtrlCreateButton("Next", 8, 136, 81, 25)
$check = GUICtrlCreateCheckbox("auto", 112, 192, 49, 17)
$Label1 = GUICtrlCreateLabel("aaaaaaaaaaaaaaaaaaaaaaaaaaaa", 8, 8, 172, 17, $SS_CENTERIMAGE)
$Label2 = GUICtrlCreateLabel("aaaaaaaaaaaaaaaaaaaaaaaaaaaa", 8, 32, 172, 17, $SS_CENTERIMAGE)
$Label3 = GUICtrlCreateLabel("aaaaaaaaaaaaaaaaaaaaaaaaaaaa", 8, 56, 172, 17, $SS_CENTERIMAGE)
$Label4 = GUICtrlCreateLabel("aaaaaaaaaaaaaaaaaaaaaaaaaaaa", 8, 80, 172, 17, $SS_CENTERIMAGE)
$Label5 = GUICtrlCreateLabel("aaaaaaaaaaaaaaaaaaaaaaaaaaaa", 8, 104, 172, 17, $SS_CENTERIMAGE)
$Label6 = GUICtrlCreateLabel("1 / 10", 32, 192, 33, 17, $SS_CENTERIMAGE)
$use = GUICtrlCreateButton("Use this one", 96, 136, 81, 25)
$list = GUICtrlCreateList("", 0, 216, 185, 56, BitOR($LBS_NOTIFY,$LBS_SORT,$WS_HSCROLL,$WS_VSCROLL))
$update = GUICtrlCreateButton("Update list", 96, 280, 83, 25)
$save = GUICtrlCreateButton("Save current", 8, 280, 83, 25)
$default = GUICtrlCreateButton("Enable DHCP", 8, 168, 81, 25)
$watchConf = GUICtrlCreateButton("Display config", 96, 168, 83, 25)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###
Global $fileName = "network.conf"
Global $confFile = FileRead($fileName)
Global $nbrOfLine = _FileCountLines($fileName)
Global $nbrOfPage = $nbrOfLine / 7
_ReloadConfig($fileName)
$lastPage = 1
$page = 1
_DisplayData(_GetData($page))
_GetListOfDevice()

While 1
   $nMsg = GUIGetMsg()
   Switch $nMsg
	  Case $GUI_EVENT_CLOSE
		 Exit
	  Case $next
		 $lastPage = $page
		 $page = $page + 1
		 if($page > $nbrOfPage) Then
			$page = 1
			$lastPage = 1
		 EndIf
		 _DisplayData(_GetData($page))
	  Case $update
		 _GetListOfDevice()
		 $data = _GetData($page)
	  Case $use
		 _SetData(_GetData($page))
	  Case $save
		 _SaveCurrentConfig($fileName)
	  Case $default
		 _ResetToDefaut()
	  Case $watchConf
		 _ActualConf()
	  EndSwitch
	  If(Mod(@SEC, 30)) Then
;~ 		 MsgBox(0,"","")
		 _ReloadConfig($fileName)
		 GUICtrlSetData($Label6,$page&" / "&$nbrOfPage)
	  EndIf
WEnd

Func _ActualConf()
   $listOfDevice = _GetNetworkAdapterInfos()
   $selectedCard = _StringBetween2(GUICtrlRead($list), "[", "]")
   If($selectedCard <> "")Then
	  $datas = (","&$listOfDevice[$selectedCard][16] &","& $listOfdevice[$selectedCard][25] &","& $listOfdevice[$selectedCard][29]&","& $listOfdevice[$selectedCard][30])
	  Local $temp[6] = _StringBetween($datas,",",",")
	  Local $dataToDisplay[6]
	  If(UBound($temp) < 6)Then
		 MsgBox(16,"Error","I can't catch all data from the selected device...")
	  Else
		 For $i=0 To 5 Step 1
			$dataToDisplay[$i] = $temp[$i]
		 Next
		 MsgBox(64,"Actual Configuration", "Interface name: "&$listOfDevice[_StringBetween2(GUICtrlRead($list), "[", "]")][7]&@CRLF&"IP ADDR: "&$dataToDisplay[3]&@CRLF&"MASQUE: "&$dataToDisplay[5]&@CRLF&"GATEWAY: "&$dataToDisplay[0]&@CRLF&"DNS1/2: "&$dataToDisplay[1]&" / "&$dataToDisplay[2])
	  EndIf
   Else
	  MsgBox(16,"Error","Please select something before...")
   EndIf
EndFunc

Func _ResetToDefaut()
   $listOfDevice = _GetNetworkAdapterInfos()
   $t = _EnableDHCP($listOfDevice[_StringBetween2(GUICtrlRead($list), "[", "]")][7])
   $y = _EnableDHCP_DNS(GUICtrlRead($list))
EndFunc

Func _ReloadConfig($configName)
   $confFile = FileRead($configName)
   $nbrOfLine = _FileCountLines($configName)
   $nbrOfPage = $nbrOfLine / 7
EndFunc

Func _SaveCurrentConfig($configName)
   $listOfDevice = _GetNetworkAdapterInfos()
   $selectedCard = _StringBetween2(GUICtrlRead($list), "[", "]")
   If($selectedCard <> "")Then
	  $datas = (","&$listOfDevice[$selectedCard][16] &","& $listOfdevice[$selectedCard][25] &","& $listOfdevice[$selectedCard][29]&","& $listOfdevice[$selectedCard][30])
	  Local $temp[6] = _StringBetween($datas,",",",")
	  Local $dataToSave[6]
	  If(UBound($temp) < 6)Then
		 MsgBox(16,"Error","I can't catch all data from the selected device...")
	  Else
		 For $i=0 To 5 Step 1
			$dataToSave[$i] = $temp[$i]
		 Next
		 $fileOpen = FileOpen($configName,1)
		 FileWriteLine($fileOpen,"[CONF"&$nbrOfPage+1&"]")
		 FileWriteLine($fileOpen,"[IP]"&$dataToSave[3]&"[EOL]")
		 FileWriteLine($fileOpen,"[MASQUE]"&$dataToSave[5]&"[EOL]")
		 FileWriteLine($fileOpen,"[GATEWAY]"&$dataToSave[0]&"[EOL]")
		 FileWriteLine($fileOpen,"[DNS1]"&$dataToSave[1]&"[EOL]")
		 FileWriteLine($fileOpen,"[DNS2]"&$dataToSave[2]&"[EOL]")
		 FileWriteLine($fileOpen,"[END"&$nbrOfPage+1&"]")
		 _ReloadConfig($configName)
	  EndIf
;~ 	  _ArrayDisplay($dataToSave)
   Else
	  MsgBox(16,"Error","Please select something before...")
   EndIf
EndFunc

Func _GetListOfDevice()
   $listOfDevice = _GetNetworkAdapterInfos()
   $counter = 0
   For $i=UBound($listOfDevice)-1 To 0 Step -1
		 GUICtrlSetData($list,"["&$i&"] "&$listOfDevice[$i][8])
		 $counter = $counter + 1
   Next
EndFunc

Func _SetData($data)
   Local $dns[2]
   $dns[0] = $data[3]
   $dns[1] = $data[4]
   _ArrayDisplay($dns)
   _SetDNSServerSearchOrder($listOfDevice[_StringBetween2(GUICtrlRead($list), "[", "]")][7],$dns)
   _EnableStatic($listOfDevice[_StringBetween2(GUICtrlRead($list), "[", "]")][7], $data[0], $data[1])
   _SetGateways($listOfDevice[_StringBetween2(GUICtrlRead($list), "[", "]")][7], $data[2])
EndFunc

Func _DisplayData($datas)
   GUICtrlSetData($Label1,"IP ADDR: "&$datas[0])
   GUICtrlSetData($Label2,"MASQUE: "&$datas[1])
   GUICtrlSetData($Label3,"GATEWAY: "&$datas[2])
   GUICtrlSetData($Label4,"DNS1: "&$datas[3])
   GUICtrlSetData($Label5,"DNS2: "&$datas[4])
   GUICtrlSetData($Label6,$page&" / "&$nbrOfPage)
EndFunc

Func _GetData($numPage)
   Local $datas[5]
   $conf = _StringBetween2($confFile, "[CONF"&$numPage&"]", "[END"&$numPage&"]")
   $datas[0] = _StringBetween2($conf, "[IP]", "[EOL]")
   $datas[1] = _StringBetween2($conf, "[MASQUE]", "[EOL]")
   $datas[2] = _StringBetween2($conf, "[GATEWAY]", "[EOL]")
   $datas[3] = _StringBetween2($conf, "[DNS1]", "[EOL]")
   $datas[4] = _StringBetween2($conf, "[DNS2]", "[EOL]")
   Return $datas
EndFunc

Func _StringBetween2($s, $from, $to) ;Help to get the string between two other. Seems to be pretty limited but still work with little strings. Don't really know where it came from.
	$x = StringInStr($s, $from) + StringLen($from)
	$y = StringInStr(StringTrimLeft($s, $x), $to)
	Return StringMid($s, $x, $y)
EndFunc   ;==>_StringBetween2
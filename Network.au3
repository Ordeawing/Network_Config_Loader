#CS
#INDEX ========================================================================================================================
Title................: Network informations and configuration for AutoIt
Filename.............: network.au3
Author...............: JGUINCH
Version..............: 3.3.14.2
Date.................: 2016-03-16

! #RequireAdmin is needed to use some functions !

Available functions :
	_DisableNetAdapter
	_EnableDHCP
	_EnableDHCP_DNS
	_EnableNetAdapter
	_EnableStatic
	_FlushDNS
	_FlushDNSEntry
	_GetNetworkAdapterFromID
	_GetNetworkAdapterInfos
	_GetNetworkAdapterList
	_GetNetworkGUI
	_GetNetworkIDFromAdapter
	_IsWirelessAdapter
	_ReleaseDHCPLease
	_RenewDHCPLease
	_SetDNSDomain
	_SetDNSServerSearchOrder
	_SetDNSSuffixSearchOrder
	_SetDynamicDNSRegistration
	_SetGateways
	_SetWINSServer

Internal functions :
	_ErrFunc
	_WMIDate
	_Array2String (from array.au3)



Remarks.......:
Some functions may fail (returns 0). @error is set to these values :
 64   = Method not supported on this platform.
 65   = Unknown failure.
 66   = Invalid subnet mask.
 67   = An error occurred while processing an instance that was returned.
 68   = Invalid input parameter.
 69   = More than five gateways specified.
 70   = Invalid IP address.
 71   = Invalid gateway IP address.
 72   = An error occurred while accessing the registry for the requested information.
 73   = Invalid domain name.
 74   = Invalid host name.
 75   = No primary or secondary WINS server defined.
 76   = Invalid file.
 77   = Invalid system path.
 78   = File copy failed.
 79   = Invalid security parameter.
 80   = Unable to configure TCP/IP service.
 81   = Unable to configure DHCP service.
 82   = Unable to renew DHCP lease.
 83   = Unable to release DHCP lease.
 84   = IP not enabled on adapter.
 85   = IPX not enabled on adapter.
 86   = Frame or network number bounds error.
 87   = Invalid frame type.
 88   = Invalid network number.
 89   = Duplicate network number.
 90   = Parameter out of bounds.
 91   = Access denied.
 92   = Out of memory.
 93   = Already exists.
 94   = Path, file, or object not found.
 95   = Unable to notify service.
 96   = Unable to notify DNS service.
 97   = Interface not configurable.
 98   = Not all DHCP leases could be released or renewed.
 100  = DHCP not enabled on the adapter.
#==============================================================================================================================

#CE


Global Const $wbemFlagReturnImmediately = 0x10
Global Const $wbemFlagForwardOnly = 0x20

Global $oErrorHandler = ObjEvent("AutoIt.Error", "_ErrFunc")
Global $errFunc = 0


; #FUNCTION# ====================================================================================================================
; Name...........: _DisableNetAdapter
; Description....: Disables the specified network adapter.
; Syntax.........: _DisableNetAdapter($sNetAdapter)
; Parameters.....: $sNetAdapter    - Name of the network adapter.
;                                    The Windows network connection name can be used instead of network adapter.
; Return values..: Success  - 1
;                  Failure  - 0
; ===============================================================================================================================
Func _DisableNetAdapter($sNetAdapter)
	Local $objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\CIMV2")
	If $objWMIService = 0 Then Return 0

	Local $sQueryNetAdapterConfig, $colNetAdapterConfig, $objNetAdapter , $iReturn = 0
	Local $adapterName

	$adapterName = _GetNetworkAdapterFromID($sNetAdapter)
	If $adapterName Then $sNetAdapter = $adapterName

	$sQueryNetAdapterConfig = "select * from Win32_NetworkAdapter Where Name = '" & $sNetAdapter & "'"

	$colNetAdapterConfig = $objWMIService.ExecQuery($sQueryNetAdapterConfig, "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	If NOT IsObj($colNetAdapterConfig) Then Return 0

	For $objNetAdapter In $colNetAdapterConfig
		$iReturn = $objNetAdapter.Disable()
		If $iReturn == 0 Then Return 1
	Next

	Return 0

EndFunc ; ===> _DisableNetAdapter

; #FUNCTION# ====================================================================================================================
; Name...........: _EnableDHCP
; Description....: Enables the Dynamic Host Configuration Protocol (DHCP) for IP configuration.
; Syntax.........: _EnableDHCP($sNetAdapter)
; Parameters.....: $sNetAdapter    - Name of the network adapter.
;                                    The Windows network connection name can be used instead of network adapter.
; Return values..: Success  - 1 or 2 (reboot required)
;                  Failure  - 0 and sets @error (see remarks)
;
; Remarks........: When the function fails (returns 0) @error contains extended information (see global remarks).
; ===============================================================================================================================
Func _EnableDHCP($sNetAdapter)
	Local $objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\CIMV2")
	If $objWMIService = 0 Then Return 0

	Local $sQueryNetAdapterConfig, $colNetAdapterConfig, $objNetAdapter, $iReturn = 0
	Local $adapterName

	$adapterName = _GetNetworkAdapterFromID($sNetAdapter)
	If $adapterName Then $sNetAdapter = $adapterName

	$sQueryNetAdapterConfig = "select * from Win32_NetworkAdapterConfiguration Where Caption like '%" & $sNetAdapter & "'"

	$colNetAdapterConfig = $objWMIService.ExecQuery($sQueryNetAdapterConfig, "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	If NOT IsObj($colNetAdapterConfig) Then Return 0

	For $objNetAdapter In $colNetAdapterConfig
		$iReturn = $objNetAdapter.EnableDHCP()

		If _EnableDHCP_DNS($sNetAdapter) Then
			If $iReturn == 0 Then return 1
			If $iReturn == 1 Then return 2
		EndIf
	Next

	If $iReturn == "" Then $iReturn = 65
	Return SetError($iReturn, "", 0)

EndFunc ; ===> _EnableDHCP

; #FUNCTION# ====================================================================================================================
; Name...........: _EnableDHCP_DNS
; Description....: Enables the Dynamic Host Configuration Protocol (DHCP) for DNS configuration.
; Syntax.........: _EnableDHCP_DNS($sNetAdapter)
; Parameters.....: $sNetAdapter    - Name of the network adapter.
;                                    The Windows network connection name can be used instead of network adapter.
; Return values..: Success  - 1
;                  Failure  - 0
; ===============================================================================================================================
Func _EnableDHCP_DNS($sNetAdapter)
	Local $objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\CIMV2")
	If $objWMIService = 0 Then Return 0

	Local $sQueryNetAdapterConfig, $colNetAdapterConfig, $objNetAdapter
	Local $guid
	Local $adapterName

	$adapterName = _GetNetworkAdapterFromID($sNetAdapter)
	If $adapterName Then $sNetAdapter = $adapterName

	$guid = _GetNetworkGUI($sNetAdapter)
	If $guid == 0 Then Return 0

	$sQueryNetAdapterConfig = "select * from Win32_NetworkAdapterConfiguration Where Caption like '%" & $sNetAdapter & "'"

	$colNetAdapterConfig = $objWMIService.ExecQuery($sQueryNetAdapterConfig, "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	If NOT IsObj($colNetAdapterConfig) Then Return 0

	For $objNetAdapter In $colNetAdapterConfig
		If $objNetAdapter.DHCPEnabled == False Then Return 0
	Next

	Return RegWrite("HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\" & $guid , "NameServer", "REG_SZ", "")

EndFunc ; ===> _EnableDHCP_DNS

; #FUNCTION# ====================================================================================================================
; Name...........: _EnableNetAdapter
; Description....: Enables the specified network adapter.
; Syntax.........: _EnableNetAdapter($sNetAdapter)
; Parameters.....: $sNetAdapter    - Name of the network adapter.
;                                    The Windows network connection name can be used instead of network adapter.
; Return values..: Success  - 1
;                  Failure  - 0
; ===============================================================================================================================
Func _EnableNetAdapter($sNetAdapter)
	Local $objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\CIMV2")
	If $objWMIService = 0 Then Return 0

	Local $sQueryNetAdapterConfig, $colNetAdapterConfig, $objNetAdapter, $iReturn = 0
	Local $adapterName

	$adapterName = _GetNetworkAdapterFromID($sNetAdapter)
	If $adapterName Then $sNetAdapter = $adapterName

	$sQueryNetAdapterConfig = "select * from Win32_NetworkAdapter Where Name = '" & $sNetAdapter & "'"

	$colNetAdapterConfig = $objWMIService.ExecQuery($sQueryNetAdapterConfig, "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	If NOT IsObj($colNetAdapterConfig) Then Return 0

	For $objNetAdapter In $colNetAdapterConfig
		$iReturn = $objNetAdapter.Enable()
		If $iReturn == 0 Then Return 1
	Next

	Return 0

EndFunc ; ===> _EnableNetAdapter

; #FUNCTION# ====================================================================================================================
; Name...........: _EnableStatic
; Description....: Enables static TCP/IP addressing for the specified network adapter.
;                  As a result, DHCP for this network adapter is disabled.
; Syntax.........: _EnableStatic($sNetAdapter, $sIPAddress, $sSubnetMask)
; Parameters.....: $sNetAdapter    - Name of the network adapter.
;                                    The Windows network connection name can be used instead of network adapter.
;                  $IPAddress      - IP addresse(s) the set to the specified network adapter. Example: 155.34.22.0.
;                                    To specify only one IP address, $IPAddress can be a string
;                                    To specify more than one IP address, $sIPAddress must be an array
;                  $sSubnetMask    - Subnet masks that complement the values in the $IPAddress parameter. Example: 255.255.0.0.
;                                    To specify more than one subnet mask, $sSubnetMask must be an array
; Return values..: Success  - 1 or 2 (reboot required)
;                  Failure  - 0 and sets @error (see remarks)
;
; Remarks........: When the function fails (returns 0) @error contains extended information (see global remarks).
; ===============================================================================================================================
Func _EnableStatic($sNetAdapter, $sIPAddress, $sSubnetMask)
	Local $objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\CIMV2")
	If $objWMIService = 0 Then Return 0

	Local $sQueryNetAdapterConfig, $colNetAdapterConfig, $objNetAdapterConfig, $iReturn = 0
	Local $aIPAddress[1]  = [ $sIPAddress  ]
	Local $aSubnetMask[1] = [ $sSubnetMask ]
	Local $adapterName

	If IsArray($sIPAddress) Then $aIPAddress = $sIPAddress
	If IsArray($sSubnetMask) Then $aSubnetMask = $sSubnetMask


	$adapterName = _GetNetworkAdapterFromID($sNetAdapter)
	If $adapterName Then $sNetAdapter = $adapterName

	$sQueryNetAdapterConfig = "select * from Win32_NetworkAdapterConfiguration Where Caption like '%" & $sNetAdapter & "'"

	$colNetAdapterConfig = $objWMIService.ExecQuery($sQueryNetAdapterConfig, "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	If NOT IsObj($colNetAdapterConfig) Then Return 0

	For $objNetAdapterConfig In $colNetAdapterConfig
		$iReturn = $objNetAdapterConfig.EnableStatic($aIPAddress, $aSubnetMask)
		If $iReturn == 0 Then Return 1
		If $iReturn == 1 Then return 2
	Next

	If $iReturn == "" Then $iReturn = 65
	Return SetError($iReturn, "", 0)

EndFunc ; ===> _EnableStatic

; #FUNCTION# ====================================================================================================================
; Name...........: _FlushDNS
; Description....: Reset the client resolver cache.
; Syntax.........: _FlushDNS
; Return values..: Success  - 1
;                  Failure  - 0
; ===============================================================================================================================
Func _FlushDNS()
	Local $aReturn
	$aReturn = DllCall("dnsapi.dll", "BOOL", "DnsFlushResolverCache")
	Return $aReturn[0]
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _FlushDNSEntry
; Description....: Remove the specified entry from the client resolver cache.
; Syntax.........: _FlushDNS
; Return values..: Success  - 1
;                  Failure  - 0
; ===============================================================================================================================
Func _FlushDNSEntry($sHost)
	Local $aReturn
	$aReturn = DllCall("dnsapi.dll", "int", "DnsFlushResolverCacheEntry_W", "WSTR", $sHost)
	If @error Then Return SetError(@error, 0, 0)
 	Return $aReturn[0]
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _GetNetworkAdapterFromID
; Description....: Get the network card name from its Windows network connection name
; Syntax.........: _GetNetworkAdapterFromID($sNetworkID)
; Parameters.....: $sNetworkID        - Name of the network connection (ex: Local Area Network).
; Return values..: Success  - Returns the network adapter name
;                  Failure  - 0
; ===============================================================================================================================
Func _GetNetworkAdapterFromID($sNetworkID)
	Local $aNetworkList = _GetNetworkAdapterList()
	If $aNetworkList = 0 Then Return 0

	For $i = 0 To UBound($aNetworkList, 1) - 1
		If $aNetworkList[$i][1] = $sNetworkID Then Return $aNetworkList[$i][0]
	Next
	Return 0
EndFunc ; ===> _GetNetConnectionID

; #FUNCTION# ====================================================================================================================
; Name...........: _GetNetworkAdapterInfos
; Description....: Retrieve informations for the specified network card.
;                  If no network adapter is specified (default), the function returns informations for all network adapters.
; Syntax.........: _GetNetworkAdapterInfos([$sNetAdapter])
; Parameters.....: $sNetAdapter        - Name of the network adapter or network ID
;                                        The Windows network connection name can be used instead of network adapter.
; Return values..: Success  - Returns a 2 dimensional array containing informations about the adapter configuration :
;                   - element[n][0] = AdapterType                    - Network adapter type.
;                       "Ethernet 802.3"
;                       "Token Ring 802.5"
;                       "Fiber Distributed Data Interface (FDDI)"
;                       "Wide Area Network (WAN)"
;                       "LocalTalk"
;                       "Ethernet using DIX header format"
;                       "ARCNET"
;                       "ARCNET (878.2)"
;                       "ATM"
;                       "Wireless"
;                       "Infrared Wireless"
;                       "Bpc"
;                       "CoWan"
;                       "1394"
;                   - element[n][1] = DeviceID                       - Unique identifier of the network adapter.
;                   - element[n][2] = GUID                           - Globally unique identifier for the connection.
;                   - element[n][3] = Index                          - Index number of the network adapter, stored in the system registry.
;                   - element[n][4] = InterfaceIndex                 - Index value that uniquely identifies the local network interface.
;                                                                      Windows XP:  This property is not available.
;                   - element[n][5] = MACAddress                     - MAC address for this network adapter.
;                   - element[n][6] = Manufacturer                   - Name of the network adapter's manufacturer.
;                   - element[n][7] = Name                           - Label by which the object is known.
;                   - element[n][8] = NetConnectionID                - Name of the network connection as it appears in the Network Connections Control Panel program.
;                   - element[n][9] = NetConnectionStatus            - State of the network adapter connection to the network.
;                       0 (0x0)  Disconnected
;                       1 (0x1)  Connecting
;                       2 (0x2)  Connected
;                       3 (0x3)  Disconnecting
;                       4 (0x4)  Hardware not present
;                       5 (0x5)  Hardware disabled
;                       6 (0x6)  Hardware malfunction
;                       7 (0x7)  Media disconnected
;                       8 (0x8)  Authenticating
;                       9 (0x9)  Authentication succeeded
;                       10 (0xA)  Authentication failed
;                       11 (0xB)  Invalid address
;                       12 (0xC)  Credentials required
;                   - element[n][10] = NetEnabled                    - Indicates whether the adapter is enabled or not.
;                   - element[n][11] = PNPDeviceID                   - Windows Plug and Play device identifier of the logical device.
;                   - element[n][12] = ProductName                   - Product name of the network adapter.
;                   - element[n][13] = ServiceName                   - Service name of the network adapter.
;                   - element[n][14] = Speed                         - Estimate of the current bandwidth in bits per second.
;                   - element[n][15] = DatabasePath                  - Valid Windows file path to standard Internet database files
;                   - element[n][16] = DefaultIPGateway              - List of IP addresses of default gateways that the computer system uses (comma-separated values).
;                   - element[n][17] = DHCPEnabled                   - If TRUE, the DHCP server automatically assigns an IP address to the computer system when establishing a network connection.
;                   - element[n][18] = DHCPLeaseExpires              - Expiration date and time for a leased IP address that was assigned to the computer by the DHCP server (YYYY-MM-DD HH:MM:SS format).
;                   - element[n][19] = DHCPLeaseObtained             - Date and time the lease was obtained for the IP address assigned to the computer by the DHCP server (YYYY-MM-DD HH:MM:SS format).
;                   - element[n][20] = DHCPServer                    - IP address of the DHCP server.
;                   - element[n][21] = DNSDomain                     - Organization name followed by a period and an extension that indicates the type of organization.
;                   - element[n][22] = DNSDomainSuffixSearchOrder    - List of DNS domain suffixes to be appended to the end of host names during name resolution (comma-separated values).
;                   - element[n][23] = DNSEnabledForWINSResolution   - If TRUE, the DNS is enabled for name resolution over WINS resolution.
;                   - element[n][24] = DNSHostName                   - Host name used to identify the local computer for authentication by some utilities.
;                   - element[n][25] = DNSServerSearchOrder          - List of server IP addresses to be used in querying for DNS servers (comma-separated values).
;                   - element[n][26] = DomainDNSRegistrationEnabled  - If TRUE, the IP addresses for this connection are registered in DNS under the domain name of this connection in addition to being registered under the computer's full DNS name.
;                   - element[n][27] = FullDNSRegistrationEnabled    - If TRUE, the IP addresses for this connection are registered in DNS under the computer's full DNS name.
;                   - element[n][28] = GatewayCostMetric             - List of integer cost metric values (ranging from 1 to 9999) to be used in calculating the fastest, most reliable, or least resource-intensive routes (comma-separated values).
;                   - element[n][29] = IPAddress                     - List of all of the IP addresses associated with the specified network adapter (comma-separated values).
;                   - element[n][30] = IPSubnet                      - List of all of the subnet masks associated with the current network adapter.
;                   - element[n][31] = WINSEnableLMHostsLookup       - If TRUE, local lookup files are used. Lookup files will contain a map of IP addresses to host names.
;                   - element[n][32] = WINSPrimaryServer             - IP address for the primary WINS server.
;                   - element[n][33] = WINSSecondaryServer           - IP address for the secondary WINS server.
;                  Failure  - 0
; ===============================================================================================================================
Func _GetNetworkAdapterInfos($sNetAdapter = "")
	Local $objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\CIMV2")
	If $objWMIService = 0 Then Return 0

	Local $aInfos[1][1], $sQueryNetAdapter, $sQueryNetAdapterConfig, $colNetAdapter, $colNetAdapterConfig, $objNetAdapter, $objNetAdapterConfig
	Local $filter = "", $n = 0, $adapterIndex
	Local $DeviceID


	If $sNetAdapter <> "" Then $filter = " WHERE Name = '" & $sNetAdapter & "' OR NetConnectionID = '" & $sNetAdapter & "'"

	$sQueryNetAdapter = 'select * from Win32_NetworkAdapter' & $filter
	$colNetAdapter = $objWMIService.ExecQuery($sQueryNetAdapter, "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)

	If NOT IsObj($colNetAdapter) Then Return 0


	For $objNetAdapter In $colNetAdapter

		$adapterIndex = $objNetAdapter.Index

		$sQueryNetAdapterConfig = 'select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True AND Index = ' & $adapterIndex
		$colNetAdapterConfig = $objWMIService.ExecQuery($sQueryNetAdapterConfig, "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)

		If IsObj($colNetAdapterConfig) Then


			For $objNetAdapterConfig In $colNetAdapterConfig
				$n += 1
				Redim $aInfos[$n][34]

				$aInfos[$n - 1][0] = $objNetAdapter.AdapterType
				$aInfos[$n - 1][1] = $objNetAdapter.DeviceID

				$DeviceID = StringFormat("%04s", $aInfos[$n - 1][1])
				$aInfos[$n - 1][2] = RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}\" & $DeviceID, "NetCfgInstanceId")

				$aInfos[$n - 1][3] = $objNetAdapter.Index
				$aInfos[$n - 1][4] = $objNetAdapter.InterfaceIndex
				$aInfos[$n - 1][5] = $objNetAdapter.MACAddress
				$aInfos[$n - 1][6] = $objNetAdapter.Manufacturer
				$aInfos[$n - 1][7] = $objNetAdapter.Name
				$aInfos[$n - 1][8] = $objNetAdapter.NetConnectionID
				$aInfos[$n - 1][9] = $objNetAdapter.NetConnectionStatus
				$aInfos[$n - 1][10] = $objNetAdapter.NetEnabled
				$aInfos[$n - 1][11] = $objNetAdapter.PNPDeviceID
				$aInfos[$n - 1][12] = $objNetAdapter.ProductName
				$aInfos[$n - 1][13] = $objNetAdapter.ServiceName
				$aInfos[$n - 1][14] = $objNetAdapter.Speed

				$aInfos[$n - 1][15]  = $objNetAdapterConfig.DatabasePath
				$aInfos[$n - 1][16]  = _Array2String( ($objNetAdapterConfig.DefaultIPGateway) , ",")
				$aInfos[$n - 1][17]  = $objNetAdapterConfig.DHCPEnabled
				$aInfos[$n - 1][18]  = _WMIDate($objNetAdapterConfig.DHCPLeaseExpires)
				$aInfos[$n - 1][19] = _WMIDate($objNetAdapterConfig.DHCPLeaseObtained)
				$aInfos[$n - 1][20] = $objNetAdapterConfig.DHCPServer
				$aInfos[$n - 1][21] = $objNetAdapterConfig.DNSDomain
				$aInfos[$n - 1][22] = _Array2String( ($objNetAdapterConfig.DNSDomainSuffixSearchOrder) , ",")
				$aInfos[$n - 1][23] = $objNetAdapterConfig.DNSEnabledForWINSResolution
				$aInfos[$n - 1][24] = $objNetAdapterConfig.DNSHostName
				$aInfos[$n - 1][25] = _Array2String( ($objNetAdapterConfig.DNSServerSearchOrder) , ",")
				$aInfos[$n - 1][26] = $objNetAdapterConfig.DomainDNSRegistrationEnabled
				$aInfos[$n - 1][27] = $objNetAdapterConfig.FullDNSRegistrationEnabled
				$aInfos[$n - 1][28] = _Array2String( ( $objNetAdapterConfig.GatewayCostMetric) , ",")
				$aInfos[$n - 1][29] = _Array2String( ( $objNetAdapterConfig.IPAddress) , ",")
				$aInfos[$n - 1][30] = _Array2String( ( $objNetAdapterConfig.IPSubnet) , ",")
				$aInfos[$n - 1][31] = $objNetAdapterConfig.WINSEnableLMHostsLookup
				$aInfos[$n - 1][32] = $objNetAdapterConfig.WINSPrimaryServer
				$aInfos[$n - 1][33] = $objNetAdapterConfig.WINSSecondaryServer
			Next
		EndIf


	Next

	If $n = 0 Then Return 0

	Return $aInfos

EndFunc ; ===> _GetNetworkAdapterInfos

; #FUNCTION# ====================================================================================================================
; Name...........: _GetNetworkAdapterList
; Description....: Get a list of Ethernet network adapter installed on the system
; Syntax.........: _GetNetworkAdapterList()
; Return values..: Success  - Returns a 2 dimensional array.
;                    element[n][0] - Name of the network adapter
;                    element[n][1] - Name of the network ID
;                  Failure  - 0
; ===============================================================================================================================
Func _GetNetworkAdapterList()
	Local $aInfos[1][2]
	Local $i = 1, $j = 1, $k = 0
	Local $aGuids[1], $sGuid, $sIndex

	While 1
		$sGuid = RegEnumKey("HKLM\SYSTEM\CurrentControlSet\services\Tcpip\Parameters\Interfaces", $i)
		If @error Then ExitLoop

		Redim $aGuids[$i]
		$aGuids[$i - 1] = $sGuid

		$i += 1
	WEnd

	$i = 1
	While 1
		$sIndex = RegEnumKey("HKLM\SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}", $i)
		If @error Then ExitLoop

		For $j = 0 To UBound($aGuids) - 1

			If RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}\" & $sIndex, "NetCfgInstanceId") = $aGuids[$j] Then
				$k += 1
				Redim $aInfos[$k][2]
				$aInfos[$k - 1][0] = RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}\" & $sIndex, "DriverDesc")
				$aInfos[$k - 1][1] = RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Network\{4D36E972-E325-11CE-BFC1-08002BE10318}\" & $aGuids[$j] & "\Connection", "Name")
			EndIf
		Next
		$i += 1
	WEnd

	If $aInfos[0][0] = "" Then Return 0
	Return $aInfos

EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _GetNetworkGUI
; Description....: Get the network ID stored in the system registry for the specified network adapter (GUID)
; Syntax.........: _GetNetworkGUI($sNetAdapter)
; Parameters.....: $sNetAdapter        - Name of the network adapter
; Return values..: Success  - Returns the network GUID
;                  Failure  - 0
; ===============================================================================================================================
Func _GetNetworkGUI($sNetAdapter)
	Local $i = 1
	Local $sIndex

	While 1
		$sIndex = RegEnumKey("HKLM\SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}", $i)
		If @error Then ExitLoop

		If RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}\" & $sIndex, "DriverDesc") = $sNetAdapter Then
			Return RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}\" & $sIndex, "NetCfgInstanceId")
		EndIf
		$i += 1
	WEnd

	Return 0
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _GetNetworkIDFromAdapter
; Description....: Get the Windows network connection name ID from the network adapter
; Syntax.........: _GetNetAdapterFromID($sNetAdapter)
; Parameters.....: $sNetAdapter        - Name of the network adapter
; Return values..: Success  - Returns the network adapter name
;                  Failure  - 0
; ===============================================================================================================================
Func _GetNetworkIDFromAdapter($sNetAdapter)
	Local $aNetworkList = _GetNetworkAdapterList()
	If $aNetworkList = 0 Then Return 0

	For $i = 0 To UBound($aNetworkList, 1) - 1
		If $aNetworkList[$i][0] = $sNetAdapter Then Return $aNetworkList[$i][1]
	Next
	Return 0
EndFunc ; ===> _GetNetConnectionID

; #FUNCTION# ====================================================================================================================
; Name...........: _IsWirelessAdapter
; Description....: Checks if a network adapter is a wireless type
; Syntax.........: _IsWirelessAdapter($sAdapter)
; Parameters.....: $sNetAdapter        - Name of the network adapter
; Return values..: Success  - Returns 1
;                  Failure  - 0
; ===============================================================================================================================
Func _IsWirelessAdapter($sAdapter)

	Local $hDLL = DllOpen("wlanapi.dll"), $aResult, $hClientHandle, $pInterfaceList, _
	$tInterfaceList, $iInterfaceCount, $tInterface, $pInterface, $tGUID, $pGUID

	$aResult = DllCall($hDLL, "dword", "WlanOpenHandle", "dword", 2, "ptr", 0, "dword*", 0, "hwnd*", 0)
	If @error Or $aResult[0] Then Return 0

	$hClientHandle = $aResult[4]

	$aResult = DllCall($hDLL, "dword", "WlanEnumInterfaces", "hwnd", $hClientHandle, "ptr", 0, "ptr*", 0)
	If @error Or $aResult[0] Then Return 0

	$pInterfaceList = $aResult[3]

	$tInterfaceList = DllStructCreate("dword", $pInterfaceList)
	$iInterfaceCount = DllStructGetData($tInterfaceList, 1)
	If Not $iInterfaceCount Then Return 0

	Local $abGUIDs[$iInterfaceCount]

	For $i = 0 To $iInterfaceCount - 1
		$pInterface = Ptr(Number($pInterfaceList) + ($i * 532 + 8))
		$tInterface = DllStructCreate("byte GUID[16]; wchar descr[256]; int State", $pInterface)
		$abGUIDs[$i] = DllStructGetData($tInterface, "GUID")

		If DllStructGetData($tInterface, "descr") == $sAdapter Then Return 1

	Next

	DllCall($hDLL, "dword", "WlanFreeMemory", "ptr", $pInterfaceList)

	$tGUID = DllStructCreate("byte[16]")
	DllStructSetData($tGUID, 1, $abGUIDs[0])
	$pGUID = DllStructGetPtr($tGUID)


	DllCall($hDLL, "dword", "WlanCloseHandle", "ptr", $hClientHandle, "ptr", 0)
	DllClose($hDLL)

	Return 0

EndFunc ; ===> _IsWirelessAdapter

; #FUNCTION# ====================================================================================================================
; Name...........: _ReleaseDHCPLease
; Description....: Releases the IP address bound to the specified (or all) DHCP-enabled network adapter.
;                  As a result, DHCP for this network adapter is disabled.
; Syntax.........: _ReleaseDHCPLease([$sNetAdapter])
; Parameters.....: $sNetAdapter    - Name of the network adapter.
;                                    If no network adapter is specified (default), the function applies to all network adapters.
;                                    The Windows network connection name can be used instead of network adapter.
; Return values..: Success  - 1 or 2 (reboot required)
;                  Failure  - 0 and sets @error (see remarks)
;
; Remarks........: When the function fails (returns 0) @error contains extended information (see global remarks).
; ===============================================================================================================================
Func _ReleaseDHCPLease($sNetAdapter = "")
	Local $objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\CIMV2")
	If $objWMIService = 0 Then Return 0

	Local  $sQueryNetAdapterConfig, $colNetAdapterConfig, $objInstance, $objNetAdapter, $iReturn = 0
	Local $adapterName

	$adapterName = _GetNetworkAdapterFromID($sNetAdapter)
	If $adapterName Then $sNetAdapter = $adapterName

	If $sNetAdapter = "" Then
		$objInstance = $objWMIService.Get("Win32_NetworkAdapterConfiguration")
		If NOT IsObj($objInstance) Then Return 0
		$iReturn = $objInstance.ReleaseDHCPLeaseAll()
	Else
		$sQueryNetAdapterConfig = "select * from Win32_NetworkAdapterConfiguration Where Caption like '%" & $sNetAdapter & "'"

		$colNetAdapterConfig = $objWMIService.ExecQuery($sQueryNetAdapterConfig, "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
		If NOT IsObj($colNetAdapterConfig) Then Return 0

		For $objNetAdapter In $colNetAdapterConfig
			$iReturn = $objNetAdapter.ReleaseDHCPLease()
			If $iReturn == 0 Then Return 1
			If $iReturn == 1 Then return 2
		Next
	EndIf

	If $iReturn == "" Then $iReturn = 65
	Return SetError($iReturn, "", 0)

EndFunc ; ===> _ReleaseDHCPLease

; #FUNCTION# ====================================================================================================================
; Name...........: _RenewDHCPLease
; Description....: Renews the IP address on the specified DHCP-enabled network adapters
; Syntax.........: _RenewDHCPLease([$sNetAdapter])
; Parameters.....: $sNetAdapter    - Name of the network adapter.
;                                    If no network adapter is specified (default), the function applies to all network adapters.
;                                    The Windows network connection name can be used instead of network adapter.
; Return values..: Success  - 1 or 2 (reboot required)
;                  Failure  - 0 and sets @error (see remarks)
;
; Remarks........: When the function fails (returns 0) @error contains extended information (see global remarks).
; ===============================================================================================================================
Func _RenewDHCPLease($sNetAdapter = "")
	Local $objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\CIMV2")
	If $objWMIService = 0 Then Return 0

	Local  $sQueryNetAdapterConfig, $colNetAdapterConfig, $objInstance, $objNetAdapter, $iReturn = 0
	Local $adapterName

	$adapterName = _GetNetworkAdapterFromID($sNetAdapter)
	If $adapterName Then $sNetAdapter = $adapterName

	If $sNetAdapter = "" Then
		$objInstance = $objWMIService.Get("Win32_NetworkAdapterConfiguration")
		If NOT IsObj($objInstance) Then Return 0
		$iReturn = $objInstance.RenewDHCPLeaseAll()
	Else
		$sQueryNetAdapterConfig = "select * from Win32_NetworkAdapterConfiguration Where Caption like '%" & $sNetAdapter & "'"

		$colNetAdapterConfig = $objWMIService.ExecQuery($sQueryNetAdapterConfig, "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
		If NOT IsObj($colNetAdapterConfig) Then Return 0

		For $objNetAdapter In $colNetAdapterConfig
			$iReturn = $objNetAdapter.RenewDHCPLease()
			If $iReturn == 0 Then Return 1
			If $iReturn == 1 Then return 2
		Next
	EndIf

	If $iReturn == "" Then $iReturn = 65
	Return SetError($iReturn, "", 0)

EndFunc ; ===> _RenewDHCPLease

; #FUNCTION# ====================================================================================================================
; Name...........: _SetDNSDomain
; Description....: Sets the DNS domain to the specified network connection/adapter
; Syntax.........: _SetDNSDomain($sNetAdapter, $sDNSDomain)
; Parameters.....: $sNetAdapter    - Name of the network adapter.
;                                    The Windows network connection name can be used instead of network adapter.
;                  $sDNSDomain     - Domain name
; Return values..: Success  - 1 or 2 (reboot required)
;                  Failure  - 0 and sets @error (see remarks)
;
; Remarks........: When the function fails (returns 0) @error contains extended information (see global remarks).
; ===============================================================================================================================
Func _SetDNSDomain($sNetAdapter, $sDNSDomain)
	Local $objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\CIMV2")
	If $objWMIService = 0 Then Return 0

	Local $sQueryNetAdapterConfig, $colNetAdapterConfig, $objNetAdapter, $iReturn = 0
	Local $adapterName

	$adapterName = _GetNetworkAdapterFromID($sNetAdapter)
	If $adapterName Then $sNetAdapter = $adapterName

	$sQueryNetAdapterConfig = "select * from Win32_NetworkAdapterConfiguration Where Caption like '%" & $sNetAdapter & "'"

	$colNetAdapterConfig = $objWMIService.ExecQuery($sQueryNetAdapterConfig, "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	If NOT IsObj($colNetAdapterConfig) Then Return 0

	For $objNetAdapter In $colNetAdapterConfig
		$iReturn = $objNetAdapter.SetDNSDomain($sDNSDomain)
		If $iReturn == 0 Then Return 1
		If $iReturn == 1 Then return 2
	Next

	If $iReturn == "" Then $iReturn = 65
	Return SetError($iReturn, "", 0)

EndFunc ; ===> _SetDNSDomain

; #FUNCTION# ====================================================================================================================
; Name...........: _SetDNSServerSearchOrder
; Description....: Sets the DNS IP addresses search order.
; Syntax.........: _SetDNSServerSearchOrder ($sNetAdapter, $aDNSServerSearchOrder)
; Parameters.....: $sNetAdapter    - Name of the network adapter.
;                                    The Windows network connection name can be used instead of network adapter.
;                  $aDNSServerSearchOrder   - (ARRAY) List of IP addresses to query for DNS servers.
; Return values..: Success  - 1 or 2 (reboot required)
;                  Failure  - 0 and sets @error (see remarks)
;
; Remarks........: When the function fails (returns 0) @error contains extended information (see global remarks).
; ===============================================================================================================================
Func _SetDNSServerSearchOrder($sNetAdapter, $aDNSServerSearchOrder)
   	Local $objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\CIMV2")
	If $objWMIService = 0 Then Return 0

	Local $sQueryNetAdapterConfig, $colNetAdapterConfig, $objNetAdapter, $iReturn = 0
	Local $adapterName
	If NOT IsArray($aDNSServerSearchOrder) Then return 0

	$adapterName = _GetNetworkAdapterFromID($sNetAdapter)
	If $adapterName Then $sNetAdapter = $adapterName

	$sQueryNetAdapterConfig = "select * from Win32_NetworkAdapterConfiguration Where Caption like '%" & $sNetAdapter & "'"

	$colNetAdapterConfig = $objWMIService.ExecQuery($sQueryNetAdapterConfig, "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	If NOT IsObj($colNetAdapterConfig) Then Return 0

	For $objNetAdapter In $colNetAdapterConfig
		$iReturn = $objNetAdapter.SetDNSServerSearchOrder($aDNSServerSearchOrder)

		If $iReturn == 0 Then Return 1
		If $iReturn == 1 Then return 2
	Next

	If $iReturn == "" Then $iReturn = 65
	Return SetError($iReturn, "", 0)

EndFunc ; ===> _SetDNSSuffixSearchOrder

; #FUNCTION# ====================================================================================================================
; Name...........: _SetDNSSuffixSearchOrder
; Description....: Sets the suffix search order.
; Syntax.........: _SetDNSSuffixSearchOrder($sNetAdapter, $aDNSDomainSuffixSearchOrder)
; Parameters.....: $sNetAdapter                     - Name of the network adapter.
;                                                     The Windows network connection name can be used instead of network adapter.
;                  $aDNSDomainSuffixSearchOrder     - (ARRAY) List of server suffixes to query for DNS servers.
; Return values..: Success  - 1 or 2 (reboot required)
;                  Failure  - 0 and sets @error (see remarks)
;
; Remarks........: When the function fails (returns 0) @error contains extended information (see global remarks).
; ===============================================================================================================================
Func _SetDNSSuffixSearchOrder($sNetAdapter, $aDNSDomainSuffixSearchOrder)
	Local $objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\CIMV2")
	If $objWMIService = 0 Then Return 0

	Local $sQueryNetAdapterConfig, $colNetAdapterConfig, $objNetAdapter, $iReturn = 0
	Local $adapterName
	If NOT IsArray($aDNSDomainSuffixSearchOrder) Then return 0

	$adapterName = _GetNetworkAdapterFromID($sNetAdapter)
	If $adapterName Then $sNetAdapter = $adapterName

	$sQueryNetAdapterConfig = "select * from Win32_NetworkAdapterConfiguration Where Caption like '%" & $sNetAdapter & "'"

	$colNetAdapterConfig = $objWMIService.ExecQuery($sQueryNetAdapterConfig, "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	If NOT IsObj($colNetAdapterConfig) Then Return 0

	For $objNetAdapter In $colNetAdapterConfig
		$iReturn = $objNetAdapter.SetDNSSuffixSearchOrder($aDNSDomainSuffixSearchOrder)
		If $iReturn == 0 Then Return 1
		If $iReturn == 1 Then return 2
	Next

	If $iReturn == "" Then $iReturn = 65
	Return SetError($iReturn, "", 0)

EndFunc ; ===> _SetDNSSuffixSearchOrder

; #FUNCTION# ====================================================================================================================
; Name...........: _SetDynamicDNSRegistration
; Description....: Indicates the dynamic DNS registration of IP addresses for the specified IP-bound adapter..
; Syntax.........: _SetDynamicDNSRegistration($sNetAdapter, $FullDNSRegistrationEnabled, $DomainDNSRegistrationEnabled)
; Parameters.....: $sNetAdapter                         - Name of the network adapter.
;                                                         The Windows network connection name can be used instead of network adapter.
;                  $FullDNSRegistrationEnabled          - If true, the IP addresses for this connection is registered in DNS
;                                                         under the computer's full DNS name.
;                  $DomainDNSRegistrationEnabled        - If true, the IP addresses for this connection are registered under the
;                                                         domain name of this connection, in addition to being registered under
;                                                         the computer's full DNS name.
; Return values..: Success  - 1 or 2 (reboot required)
;
;                  Failure  - 0 and sets @error (see remarks)
;
; Remarks........: When the function fails (returns 0) @error contains extended information (see global remarks).
; ===============================================================================================================================
Func _SetDynamicDNSRegistration($sNetAdapter, $FullDNSRegistrationEnabled, $DomainDNSRegistrationEnabled)
	Local $objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\CIMV2")
	If $objWMIService = 0 Then Return 0

	Local $sQueryNetAdapterConfig, $colNetAdapterConfig, $objNetAdapter, $iReturn = 0
	Local $adapterName

	$adapterName = _GetNetworkAdapterFromID($sNetAdapter)
	If $adapterName Then $sNetAdapter = $adapterName

	$sQueryNetAdapterConfig = "select * from Win32_NetworkAdapterConfiguration Where Caption like '%" & $sNetAdapter & "'"

	$colNetAdapterConfig = $objWMIService.ExecQuery($sQueryNetAdapterConfig, "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	If NOT IsObj($colNetAdapterConfig) Then Return 0

	For $objNetAdapter In $colNetAdapterConfig
		$iReturn = $objNetAdapter.SetDynamicDNSRegistration($FullDNSRegistrationEnabled, $DomainDNSRegistrationEnabled)
		If $iReturn == 0 Then Return 1
		If $iReturn == 1 Then return 2
	Next

	If $iReturn == "" Then $iReturn = 65
	Return SetError($iReturn, "", 0)

EndFunc ; ===> _SetDynamicDNSRegistration

; #FUNCTION# ====================================================================================================================
; Name...........: _SetGateways
; Description....: Set one or more IP gateway adresses to use for the specified network adapter
; Syntax.........: _SetGateways($sNetAdapter, $DefaultIPGateway [, $GatewayCostMetric])
; Parameters.....: $sNetAdapter        - Name of the network adapter.
;                                        The Windows network connection name can be used instead of network adapter.
;                  $DefaultIPGateway   - IP address of one or more gateways.
;                                        To specify only one gateway, $DefaultIPGateway can be a string.
;                                        To specify more than one gateway, $DefaultIPGateway must be an array.
;                  $GatewayCostMetric  - Assigns a value that ranges from 1 to 9999, which is used to calculate the fastest and
;                                        most reliable routes. Default is 1.
;                                        If multiple gateways are defined, metrics must be defined for each one in an array.
; Return values..: Success  - 1 or 2 (reboot required)
;                  Failure  - 0 and sets @error (see remarks)
;
; Remarks........: When the function fails (returns 0) @error contains extended information (see global remarks).
; ===============================================================================================================================
Func _SetGateways($sNetAdapter, $DefaultIPGateway, $GatewayCostMetric = 1)
	Local $objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\CIMV2")
	If $objWMIService = 0 Then Return 0

	Local $sQueryNetAdapterConfig, $colNetAdapterConfig, $objNetAdapterConfig, $iReturn = 0
	Local $aGWAddress[1]
	Local $aGWMetric[1]
	Local $adapterName

	$adapterName = _GetNetworkAdapterFromID($sNetAdapter)
	If $adapterName Then $sNetAdapter = $adapterName

	If IsArray($DefaultIPGateway) AND IsArray($GatewayCostMetric) Then
		$aGWAddress  = $DefaultIPGateway
		$aGWMetric  = $GatewayCostMetric
	Else
		If NOT IsArray($DefaultIPGateway) Then $aGWAddress[0] = $DefaultIPGateway
		If NOT IsArray($GatewayCostMetric) Then $aGWMetric[0] = $GatewayCostMetric
	EndIf

	$sQueryNetAdapterConfig = "select * from Win32_NetworkAdapterConfiguration Where Caption like '%" & $sNetAdapter & "'"

	$colNetAdapterConfig = $objWMIService.ExecQuery($sQueryNetAdapterConfig, "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	If NOT IsObj($colNetAdapterConfig) Then return 0

	For $objNetAdapterConfig In $colNetAdapterConfig
		$iReturn = $objNetAdapterConfig.SetGateways($aGWAddress, $aGWMetric)

		If $iReturn == 0 Then Return 1
		If $iReturn == 1 Then return 2
	Next

	If $iReturn == "" Then $iReturn = 65
	Return SetError($iReturn, "", 0)

EndFunc ; ===> _SetIPAddress

; #FUNCTION# ====================================================================================================================
; Name...........: _SetWINSServer
; Description....: Sets the primary and secondary WINS servers on the specified TCP/IP-bound network adapter
; Syntax.........: _SetWINSServer($sNetAdapter, $sWinsAddr1, $sWinsAddr2 = "")
; Parameters.....: $sNetAdapter       - Name of the network adapter.
;                                       The Windows network connection name can be used instead of network adapter.
;                  $sWinsAddr1        - IP address of the primary WINS server..
;                  $sWinsAddr2        - IP address of the secondary WINS server.
; Return values..: Success  - 1 or 2 (reboot required)
;                  Failure  - 0 and sets @error (see remarks)
;
; Remarks........: When the function fails (returns 0) @error contains extended information (see global remarks).
; ===============================================================================================================================
Func _SetWINSServer($sNetAdapter, $sWinsAddr1, $sWinsAddr2 = "")
	Local $objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\CIMV2")
	If $objWMIService = 0 Then Return 0

	Local $sQueryNetAdapterConfig, $colNetAdapterConfig, $objNetAdapter, $iReturn = 0
	Local $adapterName

	$adapterName = _GetNetworkAdapterFromID($sNetAdapter)
	If $adapterName Then $sNetAdapter = $adapterName

	$sQueryNetAdapterConfig = "select * from Win32_NetworkAdapterConfiguration Where Caption like '%" & $sNetAdapter & "'"

	$colNetAdapterConfig = $objWMIService.ExecQuery($sQueryNetAdapterConfig, "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	If NOT IsObj($colNetAdapterConfig) Then return 0

	For $objNetAdapter In $colNetAdapterConfig
		$iReturn = $objNetAdapter.SetWINSServer($sWinsAddr1, $sWinsAddr2)
		If $iReturn == 0 Then Return 1
		If $iReturn == 1 Then return 2
	Next

	If $iReturn == "" Then $iReturn = 65
	Return SetError($iReturn, "", 0)

EndFunc ; ===> _SetWINSServer

; #FUNCTION# ====================================================================================================================
; Name...........: _ErrFunc
; Description....: Empty function witch intercepts AutoIt errors
; ===============================================================================================================================
Func _ErrFunc($oError)
EndFunc   ;==>_ErrFunc




; #FUNCTION# ====================================================================================================================
; Name...........: _WMIDate
; Description....: Converts a WMI date to a standard date-time format.
; Syntax.........: _WMIDate($sDate, [$iFormat])
; Parameters.....: $sDate          - WMI date (obtnained from a WMI object)
;                  $iFormat        - Format of date that the user want to get.
;                   0 = YYYY-MM-DD HH:MM:SS ( default)
;                   1 = DD/MM/YYYY HH:MM:SS
;                   2 = YYYYMMDDHHMMSS
; Return values..: Success  - Converted date
;                  Failure  - ""
; ===============================================================================================================================
Func _WMIDate($sDate, $iFormat = 0)
	Local $dateFormat, $newdate = ""

	If $iFormat = 0 Then $dateFormat = "$1-$2-$3 $4:$5:$6"
	If $iFormat = 1 Then $dateFormat = "$3/$2/$1 $4:$5:$6"
	If $iFormat = 2 Then $dateFormat = "$1$2$3$4$5$6"


	If NOT IsString($sDate) Then Return ""

	$newDate = StringRegExpReplace($sDate, "\A(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2}).*", $dateFormat)
	Return $newDate
EndFunc ; ===> _WMIDate




; #FUNCTION# ====================================================================================================================
; Name...........: _Array2String
; Description ...: Places the elements of an array into a single string, separated by the specified delimiter.
; Syntax.........: _Array2String(Const ByRef $avArray[, $sDelim = "|"[, $iStart = 0[, $iEnd = 0]]])
; Parameters ....: $avArray - Array to combine
;                  $sDelim  - [optional] Delimiter for combined string
;                  $iStart  - [optional] Index of array to start combining at
;                  $iEnd    - [optional] Index of array to stop combining at
; Return values .: Success - string which combined selected elements separated by the delimiter string.
;                  Failure - "", sets @error:
;                  |1 - $avArray is not an array
;                  |2 - $iStart is greater than $iEnd
;                  |3 - $avArray is not an 1 dimensional array
; Author ........: Brian Keene <brian_keene at yahoo dot com>, Valik - rewritten
; Modified.......: Ultima - code cleanup
; ===============================================================================================================================
Func _Array2String($avArray, $sDelim = "|", $iStart = 0, $iEnd = 0)
	If Not IsArray($avArray) Then Return SetError(1, 0, "")
	If UBound($avArray, 0) <> 1 Then Return SetError(3, 0, "")

	Local $sResult, $iUBound = UBound($avArray) - 1

	; Bounds checking
	If $iEnd < 1 Or $iEnd > $iUBound Then $iEnd = $iUBound
	If $iStart < 0 Then $iStart = 0
	If $iStart > $iEnd Then Return SetError(2, 0, "")

	; Combine
	For $i = $iStart To $iEnd
		$sResult &= $avArray[$i] & $sDelim
	Next

	Return StringTrimRight($sResult, StringLen($sDelim))
EndFunc   ;==>_Array2String
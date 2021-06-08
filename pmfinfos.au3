#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=static\icon.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Description=pmfinfos.exe
#AutoIt3Wrapper_Res_Fileversion=1.0.4.0
#AutoIt3Wrapper_Res_ProductName=pmfinfos
#AutoIt3Wrapper_Res_ProductVersion=1.0.4
#AutoIt3Wrapper_Res_CompanyName=CNAMTS/CPAM_ARTOIS/BEI
#AutoIt3Wrapper_Res_LegalCopyright=bei.cpam-artois@assurance-maladie.fr
#AutoIt3Wrapper_Res_Language=1036
#AutoIt3Wrapper_Res_Compatibility=Win7
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=Y
#AutoIt3Wrapper_AU3Check_Parameters=-q -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#Au3Stripper_Parameters=/MO /RSLN
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

; #INDEX# =======================================================================================================================
; Title .........: pmfinfos
; AutoIt Version : 3.3.14.5
; Language ......: French
; Description ...: Script .au3
; Author(s) .....: yann.daniel@assurance-maladie.fr
; ===============================================================================================================================

; #ENVIRONMENT# =================================================================================================================
; AutoIt3Wrapper
; Includes YD
#include "C:\Users\DANIEL-03598\Autoit_dev\Include\YDGVars.au3"
#include "C:\Users\DANIEL-03598\Autoit_dev\Include\YDLogger.au3"
#include "C:\Users\DANIEL-03598\Autoit_dev\Include\YDTool.au3"
; Includes Constants
#include <StaticConstants.au3>
#Include <WindowsConstants.au3>
#include <TrayConstants.au3>
; Includes
#include <String.au3>
; Options
AutoItSetOption("MustDeclareVars", 1)
AutoItSetOption("WinTitleMatchMode", 2)
AutoItSetOption("WinDetectHiddenText", 1)
AutoItSetOption("MouseCoordMode", 0)
AutoItSetOption("TrayMenuMode", 3)
OnAutoItExitRegister("_YDTool_ExitApp")
; ===============================================================================================================================

; #VARIABLES# ===================================================================================================================
_YDGVars_Set("sAppName", _YDTool_GetAppWrapperRes("ProductName"))
_YDGVars_Set("sAppDesc", _YDTool_GetAppWrapperRes("Description"))
_YDGVars_Set("sAppVersion", _YDTool_GetAppWrapperRes("ProductVersion"))
_YDGVars_Set("sAppContact", _YDTool_GetAppWrapperRes("LegalCopyright"))
_YDGVars_Set("sAppVersionV", "v" & _YDGVars_Get("sAppVersion"))
_YDGVars_Set("sAppTitle", _YDGVars_Get("sAppName") & " - " & _YDGVars_Get("sAppVersionV"))
_YDGVars_Set("sAppDirDataPath", @ScriptDir & "\data")
_YDGVars_Set("sAppDirStaticPath", @ScriptDir & "\static")
_YDGVars_Set("sAppDirLogsPath", @ScriptDir & "\logs")
_YDGVars_Set("sAppDirVendorPath", @ScriptDir & "\vendor")
_YDGVars_Set("sAppIconPath", @ScriptDir & "\static\icon.ico")
_YDGVars_Set("sAppConfFile", @ScriptDir & "\conf.ini")
_YDGVars_Set("iAppNbDaysToKeepLogFiles", 15)

_YDLogger_Init()
_YDLogger_LogAllGVars()
; ===============================================================================================================================

; #MAIN SCRIPT# =================================================================================================================
If Not _YDTool_IsSingleton() Then Exit
;------------------------------
; On supprime les anciens fichiers de log
_YDTool_DeleteOldFiles(_YDGVars_Get("sAppDirLogsPath"), _YDGVars_Get("iAppNbDaysToKeepLogFiles"))
;------------------------------
; On cree le repertoire data s il n existe pas
_YDTool_CreateFolderIfNotExist(_YDGVars_Get("sAppDirDataPath"))
;------------------------------
; On recupere les valeurs de conf.ini
Global $g_sSiteIniFilePath = _YDGVars_Get("sAppDirDataPath") & "\site.ini"
Global $g_bConnectProgres = (_YDTool_GetAppConfValue("general", "connect_progres") = 1) ? True : False
Global $g_sSiteNumND = @ScriptDir & "\" & _YDTool_GetAppConfValue("general", "site_non_determine")
Global $g_sIPND = @ScriptDir & "\" & _YDTool_GetAppConfValue("general", "ip_non_determine")
Global $g_aConfDNS = _YDTool_GetAppConfSection("dns")
Global $g_aConfSites = _YDTool_GetAppConfSection("sites")
Global $g_aConfSitesLetter = _YDTool_GetAppConfSection("sites_letter")
Global $g_aConfSitesDHCP = _YDTool_GetAppConfSection("sites_dhcp")
;------------------------------
; On gere l'affichage de l'icone dans le tray
TraySetIcon(_YDGVars_Get("sAppIconPath"))
TraySetToolTip(_YDGVars_Get("sAppTitle"))
Global $idTrayCopyInfos = TrayCreateItem("Copier les infos du PMF", -1, -1, -1)
TrayCreateItem("")
Global $idTrayMenuSite = TrayCreateMenu("Changer de site", -1, -1)
If $g_bConnectProgres Then
	Global $idTrayReconnectProgres = TrayCreateItem("Reconnecter vos lecteurs PROGRES", -1, -1, -1)
EndIf
Global $oTraySites = ObjCreate("Scripting.Dictionary")
; On boucle sur la section [sites] de la conf et on ajoute au dictionnaire des id du TrayMenu
For $i = 1 to $g_aConfSites[0][0]
	Global $idTraySite = TrayCreateItem($g_aConfSites[$i][1], $idTrayMenuSite)
	$oTraySites.Add($idTraySite, $g_aConfSites[$i][0])
Next
TrayCreateItem("")
Global $idTrayHelp = TrayCreateItem("Aide", -1, -1, -1)
Global $idTrayAbout = TrayCreateItem("A propos", -1, -1, -1)
TrayCreateItem("")
Global $idTrayExit = TrayCreateItem("Quitter", -1, -1, -1)
TraySetState($TRAY_ICONSTATE_SHOW)
;------------------------------
; On recupere d autres variables globales
Global $g_sLoggedUserName = _YDTool_GetHostLoggedUserName(@ComputerName)
_YDLogger_Var("$g_sLoggedUserName", $g_sLoggedUserName)
Global $g_sOSArchitecture = (@OSVersion = "WIN_7") ? "x86" : "x64"
_YDLogger_Var("$g_sOSArchitecture", $g_sOSArchitecture)
Global $g_sOS = "Windows " & ($g_sOSArchitecture=="x86" ? "7" : "10")
_YDLogger_Var("$g_sOS", $g_sOS)
Global $g_sMsgNotConnectedToRamage = "/!\ Non connecte a RAMAGE /!\"
Global $g_sIP = $g_sIPND, $g_sOldIP = "255.255.255.255", $g_sRamageIPStart = "55.", $g_sMAC = "", $g_sDNSDomain = "", $g_sDNSType = "", $g_sSiteNum = "", $g_sSiteName = ""
Global $g_sHelpFilePath = @ScriptDir & "\" & _YDGVars_Get("sAppName") & ".pdf"
Global $aTraySitesKeys = $oTraySites.Keys
Global $aTraySitesItems = $oTraySites.Items
Global $iTraySitesCount = $oTraySites.Count - 1
Global $iTraySiteFirstId = Number($aTraySitesKeys[0])
Global $iTraySiteLastId = Number($aTraySitesKeys[$iTraySitesCount])
; On appelle une premiere fois la fonction de recuperation des infos reseau, puis on les rafraichi toutes les 5 secondes
_GetLocalNetworkInfo()
AdlibRegister("_GetLocalNetworkInfo", 2000)
; #MAIN SCRIPT# =================================================================================================================

; #MAIN LOOP# ====================================================================================================================
While 1
	Global $iMsg = TrayGetMsg()
	Switch $iMsg
		Case $idTrayExit
			_YDTool_ExitConfirm()
		Case $idTrayAbout
			_YDTool_GUIShowAbout()
		Case $idTrayHelp
			Run("RunDLL32.EXE url.dll,FileProtocolHandler " & $g_sHelpFilePath)
		Case $idTrayCopyInfos
			_CopyInfos()
		Case $iTraySiteFirstId To $iTraySiteLastId
			_ChangeSite($iMsg)
			If $g_bConnectProgres Then
				_ConnectProgresNetworkDrives($g_sSiteNum)
			Endif
		Case $idTrayReconnectProgres
			_ConnectProgresNetworkDrives($g_sSiteNum)
		Case Else
			_Main()
	EndSwitch
	;------------------------------
	Sleep(10)
WEnd
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Description ...: Traitement principal
; Syntax ........: _Main()
; Parameters ....:
; Return values .:
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 05/03/2021
; Notes .........:
;================================================================================================================================
Func _Main()
	Local $sFuncName = "_Main"
	; NON connecte a RAMAGE
	If _IsConnectedToRamage($g_sIP) = False Then
		If StringLeft($g_sOldIP, 3) = $g_sRamageIPStart Then
			_YDLogger_Log("Deconnecte du reseau RAMAGE !!! :-(", $sFuncName)
			$g_sOldIP = $g_sIP
		Endif
		$g_sSiteName = $g_sMsgNotConnectedToRamage
		_RefreshTrayTip()
	; Connecte a RAMAGE
	ElseIf $g_sIP <> $g_sOldIP Then
		_YDLogger_Log("Connecte au reseau RAMAGE !!! :-)", $sFuncName)
		; On sauvegarde l IP en cours dans $g_sOldIP
		_YDLogger_Var("$g_sIP", $g_sIP, $sFuncName)
		$g_sOldIP = $g_sIP
		_YDLogger_Var("$g_sOldIP", $g_sOldIP, $sFuncName, 2)
		; On recupere le numero du site
		$g_sSiteNum = _GetSiteNum($g_sIP)
		_YDLogger_Var("$g_sSiteNum", $g_sSiteNum, $sFuncName)
		; On recupere le nom du site
		$g_sSiteName = _GetSiteName($g_sSiteNum)
		_YDLogger_Var("$g_sSiteName", $g_sSiteName, $sFuncName)
		; On lance la connection des lecteurs progres si configure
		If $g_bConnectProgres Then
			Sleep(3000)
			_ConnectProgresNetworkDrives($g_sSiteNum)
		Endif
		_RefreshTrayTip()
	EndIf	
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Recuperation des infos locales (via WMIC)
; Syntax ........: _GetLocalNetworkInfo()
; Parameters ....:
; Return values .:
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 08/04/2021
; Notes .........:
;================================================================================================================================
Func _GetLocalNetworkInfo()
	Local $sFuncName = "_GetLocalNetworkInfo"
	Local $oWMIService = ObjGet('winmgmts:{impersonationLevel = impersonate}!\\' & '.' & '\root\cimv2')
	Local $oColItems = $oWMIService.ExecQuery('Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True', 'WQL', 0x30)
	Local $bFind = False
	Local $j = 0	
	; On verifie que le tableau des DNS est OK
	If Not IsArray($g_aConfDNS) Then
		_YDLogger_Error("$g_aConfDNS n'est pas un tableau !")
		Return
	EndIf
	If IsObj($oColItems) = 1 Then
		; On boucle sur les items WMIC
		For $oObjectItem In $oColItems
			If $bFind Then ExitLoop
			_YDLogger_Log(">>> Boucle $oObjectItem : " & $oObjectItem.IPAddress(0), $sFuncName, 2)
			; On boucle sur les dns autorises dans la conf
			For $i = 1 to $g_aConfDNS[0][0]
				If $bFind Then ExitLoop
				_YDLogger_Sep(50, "-", 2)
				_YDLogger_Log(">>> Boucle $g_aConfDNS : " & $i, $sFuncName, 2)
				_YDLogger_Var("$oObjectItem.IPAddress(0)", $oObjectItem.IPAddress(0), $sFuncName, 2)
				_YDLogger_Var("$oObjectItem.MACAddress", $oObjectItem.MACAddress, $sFuncName, 2)
				_YDLogger_Var("$oObjectItem.DNSDomain", $oObjectItem.DNSDomain, $sFuncName, 2)
				_YDLogger_Var("$g_aConfDNS[" & $i & "][0]", $g_aConfDNS[$i][0], $sFuncName, 2)
				_YDLogger_Var("$g_aConfDNS[" & $i & "][1]", $g_aConfDNS[$i][1], $sFuncName, 2)
				If $oObjectItem.DNSDomain = $g_aConfDNS[$i][1] Then
					$g_sIP = $oObjectItem.IPAddress(0)
					$g_sMAC = $oObjectItem.MACAddress
					$g_sDNSDomain = $oObjectItem.DNSDomain
					$g_sDNSType = $g_aConfDNS[$i][0]
					$j = $j + 1
					_YDLogger_Log("EGALITE TROUVEE !! : " & $oObjectItem.DNSDomain & " = " & $g_aConfDNS[$i][1], $sFuncName, 2)
					$bFind = True
				Else
					$g_sIP = $oObjectItem.IPAddress(0)
					$j = $j + 1
				EndIf
			Next
		Next
	EndIf
	; Si connexion non liee aux DNS attendus, on fige les valeurs en non determinee
	If $j = 0 Then
		$g_sIP = $g_sIPND
		$g_sDNSType = ""
	EndIf
	_YDLogger_Sep(50, "-", 2)
	_YDLogger_Var("$g_sIP", $g_sIP, $sFuncName, 2)
	_YDLogger_Var("$g_sMAC", $g_sMAC, $sFuncName, 2)
	_YDLogger_Var("$g_sDNSDomain", $g_sDNSDomain, $sFuncName, 2)
	_YDLogger_Var("$g_sDNSType", $g_sDNSType, $sFuncName, 2)
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Recuperation du numero du site
; Syntax ........: _GetSiteNum($g_sIP)
; Parameters ....: $_sIP 		- Adresse IP
; Return values .: $sSiteNum	- Numero du site au format SITExx
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 05/03/2021
; Notes .........:
;================================================================================================================================
Func _GetSiteNum($_sIP)
	Local $sFuncName = "_GetSiteNum"
	Local $aIP = StringSplit($_sIP, ".")
	Local $sSiteNum = $g_sSiteNumND, $sDHCPStart, $sDHCPStop
	Local $aDHCPRange, $aDHCPStart, $aDHCPStop

	_YDLogger_Var("$g_sDNSType", $g_sDNSType, $sFuncName)
	; On recherche d'abord a partir de l'IP si connecte en local	
	If $g_sDNSType = "local" Then
		_YDLogger_Log("Recherche par rapport a IP locale ...", $sFuncName)
		; On boucle sur la section [sites_dhcp] dans la conf
		For $i = 1 to $g_aConfSitesDHCP[0][0]
			_YDLogger_Var("$g_aConfSitesDHCP[$i][0]", $g_aConfSitesDHCP[$i][0], $sFuncName, 2)
			$aDHCPRange = StringSplit($g_aConfSitesDHCP[$i][1], "|")
			$sDHCPStart = $aDHCPRange[1]
			$sDHCPStop = $aDHCPRange[2]
			_YDLogger_Var("$sDHCPStart", $sDHCPStart, $sFuncName, 2)
			_YDLogger_Var("$sDHCPStop", $sDHCPStop, $sFuncName, 2)
			$aDHCPStart = StringSplit($sDHCPStart, ".")
			$aDHCPStop = StringSplit($sDHCPStop, ".")
			If Int($aIP[1]) = $aDHCPStart[1] And Int($aIP[2]) = $aDHCPStart[2] And Int($aIP[3]) >= $aDHCPStart[3] And Int($aIP[3]) <= $aDHCPStop[3] Then
				$sSiteNum = $g_aConfSitesDHCP[$i][0]
				ExitLoop
			EndIf
		Next
	Else
		; On verifie ensuite si le numero du site est deja reference dans le fichier site.ini
		If FileExists($g_sSiteIniFilePath) Then		 
			; On boucle sur la section [sites] de la conf pour verifier que le site est bien reference
			For $i = 1 to $g_aConfSites[0][0]
				_YDLogger_Var("$g_aConfSites[" & $i & "][1]", $g_aConfSites[$i][1], $sFuncName, 2)
				If $g_aConfSites[$i][0] = IniRead($g_sSiteIniFilePath, "general", "site_num", "") Then
					$sSiteNum = $g_aConfSites[$i][0]
					_YDLogger_Log("Numero de site trouve dans fichier " & $g_sSiteIniFilePath & " : " & $sSiteNum, $sFuncName)
					ExitLoop
				EndIf
			Next
		Else
			_YDLogger_Log("Fichier site.ini non trouve : " & $g_sSiteIniFilePath, $sFuncName)
		EndIf
		; Si le site n a pas ete trouve dans le fichier site.ini, recherche sur le nom du PMF
		If $sSiteNum == $g_sSiteNumND Then
			_YDLogger_Log("Recherche par rapport au nom du PMF (car VPN) ...", $sFuncName)
			; On boucle sur la section [sites_letter] dans la conf
			For $i = 1 to $g_aConfSitesLetter[0][0]
				_YDLogger_Var("$g_aConfSitesLetter[$i][1]", $g_aConfSitesLetter[$i][1], $sFuncName, 2)
				If $g_aConfSitesLetter[$i][1] = StringMid(@ComputerName, 12, 1) Then
					$sSiteNum = $g_aConfSitesLetter[$i][0]
					ExitLoop
				EndIf
			Next
		EndIf
	EndIf
	; On met a jour le fichier site.ini
	If $sSiteNum <> $g_sSiteNumND Then
		_SetSiteInIniFile($sSiteNum)
	EndIf
	; On renvoie le numero du site trouve
	_YDLogger_Var("$sSiteNum", $sSiteNum, $sFuncName)
	return $sSiteNum
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Mise a jour du numero du site dans le fichier site.ini
; Syntax ........: _SetSiteInIniFile($_sSiteNum)
; Parameters ....: $_sSiteNum	- Numero du site au format SITExx
; Return values .: True 	- Mise a jour effectue
;                  False 	- Mise a jour KO
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 09/03/2021
; Notes .........:
;================================================================================================================================
Func _SetSiteInIniFile($_sSiteNum)
	Local $sFuncName = "_SetSiteInIniFile"	
	; On boucle sur la section [sites] de la conf pour verifier que le site est bien reference
	For $i = 1 to $g_aConfSites[0][0]
		_YDLogger_Var("$g_aConfSites[$i][1]", $g_aConfSites[$i][1], $sFuncName, 2)
		; Si tout est OK on ecrit la valeur du SITExx dans le fichier site.ini
		If $g_aConfSites[$i][0] = $_sSiteNum Then
			If IniWriteSection($g_sSiteIniFilePath, "general", "site_num=" & $_sSiteNum) = 1 Then
				_YDLogger_Log("Ecriture du site dans le fichier site.ini : " & $_sSiteNum, $sFuncName)
				; On met aussi a jour les variables globales
				$g_sSiteNum = $_sSiteNum
				$g_sSiteName = _GetSiteName($_sSiteNum)
				Return True
			Else
				_YDLogger_Error("Probleme lors de la mise a jour du site dans fichier " & $g_sSiteIniFilePath & " : " & $_sSiteNum, $sFuncName)
				Return False
			EndIf
		EndIf
	Next
EndFunc	

; #FUNCTION# ====================================================================================================================
; Description ...: Recuperation du nom du site
; Syntax ........: _GetSiteName($g_sSiteNum)
; Parameters ....: $_sSiteNum	- Numero du site au format SITExx
; Return values .: $sSiteName	- Nom du site
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 05/03/2021
; Notes .........:
;================================================================================================================================
Func _GetSiteName($_sSiteNum)
	Local $sFuncName = "_GetSiteName"
	Local $sSiteName = "Non déterminé"	
	; On boucle sur la section [sites] dans la conf
	For $i = 1 to $g_aConfSites[0][0]
		_YDLogger_Var("$g_aConfSites[$i][1]", $g_aConfSites[$i][1], $sFuncName, 2)
		If $g_aConfSites[$i][0] = $_sSiteNum Then
			$sSiteName = $g_aConfSites[$i][1]
			ExitLoop
		EndIf
	Next
	_YDLogger_Var("$sSiteName", $sSiteName, $sFuncName, 2)
	return $sSiteName
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Recupere le SiteNum a partir du TraySiteId
; Syntax ........: _GetSiteNumBySiteId($_iTraySiteId)
; Parameters ....: $_iTraySiteId	- TraySiteId
; Return values .: $sSiteNum		- Numero du site au format SITExx
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 12/03/2021
; Notes .........:
;================================================================================================================================
Func _GetSiteNumBySiteId($_iTraySiteId)
	Local $sFuncName = "_GetSiteNumBySiteId"
	Local $sSiteNum = $oTraySites.Item($_iTraySiteId)
	_YDLogger_Var("$sSiteNum", $sSiteNum, $sFuncName, 2)
	return $sSiteNum
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Recupere le TraySiteId a partir du SiteNum
; Syntax ........: _GetSiteIdBySiteNum($_sSiteNum)
; Parameters ....: $_sSiteNum		- Numero du site au format SITExx
; Return values .: $iTraySiteId		- TraySiteId
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 12/03/2021
; Notes .........:
;================================================================================================================================
Func _GetSiteIdBySiteNum($_sSiteNum)
	Local $sFuncName = "_GetSiteIdBySiteNum"
	Local $iTraySiteId = 0
	For $i = 0 To $iTraySitesCount
		If $_sSiteNum = $aTraySitesItems[$i] Then
			_YDLogger_Var("TEST VAL", $aTraySitesItems[$i], $sFuncName, 2)
			_YDLogger_Var("TEST KEY", $aTraySitesKeys[$i], $sFuncName, 2)
			$iTraySiteId = $aTraySitesKeys[$i]
			_YDLogger_Var("$iTraySiteId", $iTraySiteId, $sFuncName, 2)
			ExitLoop
		EndIf
	Next
	Return $iTraySiteId
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Rafraichit les infos dans le TrayTip
; Syntax ........: _RefreshTrayTip()
; Parameters ....:
; Return values .:
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 05/03/2021
; Notes .........:
;================================================================================================================================
Func _RefreshTrayTip()
	Local $sFuncName = "_RefreshTrayTip"
	Local $sTitleForTray = ""
	; On met a jour les infos dans le TrayTip
	$sTitleForTray &= @ComputerName & @CRLF
	$sTitleForTray &= $g_sIP & @CRLF
	$sTitleForTray &= "----------" & @CRLF
	$sTitleForTray &= "Util: " & $g_sLoggedUserName & @CRLF
	$sTitleForTray &= "Site: " & $g_sSiteName & @CRLF	
	$sTitleForTray &= "Os  : " & $g_sOS & @CRLF
	; On decoche tous les sites
	For $i = 0 To $iTraySitesCount
		TrayItemSetState($aTraySitesKeys[$i], $TRAY_UNCHECKED)
	Next
	; On coche uniquement celui choisi ou affecte
	TrayItemSetState(_GetSiteIdBySiteNum($g_sSiteNum), $TRAY_CHECKED)
	_YDLogger_Log("Rafraichissement du TrayTip", $sFuncName, 2)
	TraySetToolTip($sTitleForTray)	
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Permet de copier (presse papier) les infos du PMF
; Syntax ........: _CopyInfos()
; Parameters ....:
; Return values .:
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 05/03/2021
; Notes .........:
;================================================================================================================================
Func _CopyInfos()
	Local $sFuncName = "_CopyInfos"
	Local $sInfos = @ComputerName & " - " & $g_sIP & " - " & $g_sLoggedUserName & " - " & $g_sSiteName & " - " & $g_sOS
	_YDLogger_Var("$sInfos", $sInfos, $sFuncName)
	_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Informations PMF copiees dans le presse-papier")
	ClipPut($sInfos)
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Effectue le changement de site
; Syntax ........: _ChangeSite($_iTraySiteId)
; Parameters ....: $_iTraySiteId	- Id (TrayMenu) du site selectionne
; Return values .: 
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 12/03/2021
; Notes .........:
;================================================================================================================================
Func _ChangeSite($_iTraySiteId)
	Local $sFuncName = "_ChangeSite"
	Local $sSiteNum = $g_sSiteNumND
	_YDLogger_Var("$_iTraySiteId", $_iTraySiteId, $sFuncName, 2)
	; On met a jour le fichier site.ini
	$sSiteNum = _GetSiteNumBySiteId($_iTraySiteId)
	_YDLogger_Var("$sSiteNum", $sSiteNum, $sFuncName, 2)
	If $sSiteNum <> $g_sSiteNumND Then
		_YDLogger_Log("Changement de site : " & $g_sSiteNum & " => " & $sSiteNum, $sFuncName)
		_SetSiteInIniFile($sSiteNum)
		_RefreshTrayTip()
		_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Changement de site vers " &  _GetSiteName($sSiteNum) & " : En cours ...")
	Else
		_YDLogger_Log("Site non determine", $sFuncName)
	EndIf	
Endfunc

; #FUNCTION# ====================================================================================================================
; Description ...: Permet de connecter les lecteurs PROGRES du site passe en parametre
; Syntax ........: _ConnectProgresNetworkDrives($_sSiteNum)
; Parameters ....: $_sSiteNum		- Numero du site au format SITExx
; Return values .: True 	- Tous les lecteurs sont bien connectes
;                  False 	- Tous les lecteurs ne sont pas connectes
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 17/03/2021
; Notes .........:
;================================================================================================================================
Func _ConnectProgresNetworkDrives($_sSiteNum)
	Local $sFuncName = "_ConnectProgresNetworkDrives"
	Local $aNetworkDriveLetters[3] = ["H", "M", "R"]
	Local $sNetworkPath, $sNetworkDriveLetter, $bIsAllConnected
	; On affiche un warning si le site est non determine
	If $_sSiteNum = $g_sSiteNumND Then
		_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Site non déterminé : impossible de configurer vos lecteurs PROGRES !" & @CRLF & "==> Clic sur l'icône " & _YDGVars_Get("sAppName") & " > Changer de site", 10000, 2)
		Return False
	EndIf	
	; On boucle sur la section [sites] de la conf pour verifier que le site est bien reference
	For $i = 1 to $g_aConfSites[0][0]
		_YDLogger_Var("$g_aConfSites[$i][1]", $g_aConfSites[$i][1], $sFuncName, 2)
		; Si tout est OK on va recuperer les valeurs chemins reseau
		If $g_aConfSites[$i][0] = $_sSiteNum Then
			$bIsAllConnected = True
			; On boucle sur chaque lettre du tableau des lecteurs reseaux
			For $j = 0 To Ubound($aNetworkDriveLetters) -1
				$sNetworkDriveLetter = $aNetworkDriveLetters[$j]
				_YDLogger_Var("$sNetworkDriveLetter", $sNetworkDriveLetter, $sFuncName)
				; On va chercher le chemin selon le SiteNum et la lettre du lecteur reseau
				$sNetworkPath = _YDTool_GetAppConfValue("progres_path",  $_sSiteNum & "_" & $sNetworkDriveLetter)
				_YDLogger_Var("$sNetworkPath", $sNetworkPath, $sFuncName, 2)
				; on supprime le lecteur reseau s il existe
				_YDTool_DeleteNetworkConnection($sNetworkDriveLetter & ":")				
				; On cree le lecteur reseau
				If Not _YDTool_CreateNetworkConnection($aNetworkDriveLetters[$j] & ":", $sNetworkPath) Then
					$bIsAllConnected = False
				EndIf
			Next
			_YDLogger_Var('$bIsAllConnected', $bIsAllConnected, $sFuncName, 2)
			If $bIsAllConnected Then
				_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Les lecteurs PROGRES ont ete configures selon votre site : " & $g_sSiteName)
			Else
				_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Les lecteurs PROGRES n'ont PAS tous pu etre configures selon votre site : " & $g_sSiteName, 3000, 2)
			EndIf
			ExitLoop
		EndIf
	Next
	; On liste les infos des lecteurs reseau
	_YDTool_GetNetworkConnections()	
Endfunc

; #FUNCTION# ====================================================================================================================
; Description ...: Vérification que l'IP correspond a RAMAGE
; Syntax ........: _IsConnectedToRamage($g_sIP)
; Parameters ....:
; Return values .: True 	- Connecté a RAMAGE
;                  False 	- Non connecté a RAMAGE
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 05/03/2021
; Notes .........:
;================================================================================================================================
Func _IsConnectedToRamage($_sIP)
	Local $sFuncName = "_IsConnectedToRamage"
	If StringLeft($_sIP, 3) = $g_sRamageIPStart Then
		Return True
	EndIf
	_YDLogger_Log("/!\ Non connecte a RAMAGE /!\ [IP: " & $_sIP & "]", $sFuncName, 2)
	Return False
EndFunc
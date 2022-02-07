Attribute VB_Name = "add2GinBB"
Public Cntatnd As Integer
Public CntGsmSec As Integer
Public sname As String
Public freqb As String
Public fileName As String
Public sctr As Integer
Public path As String
Public fs, D As String, fileNamedoc As String
Sub autfi()
    Sheet2.Columns.autoFit
End Sub
Sub getData2GinBB(control As IRibbonControl)
TurnOnSpeed True
Call cpFrn
Call cpAtnd
ThisWorkbook.Worksheets("SC2GBB").Columns.autoFit
ThisWorkbook.Worksheets("SC2GBB").Rows.autoFit
ThisWorkbook.Worksheets("SC2GBB").Activate
TurnOnSpeed False

End Sub
Sub cpAtnd()
    Dim jmlBrAtnd As Integer
    Dim atnd As Worksheet
    Dim shSc2Gbb As Worksheet
    Dim shFront As Worksheet
    Dim i As Integer
    Set atnd = ThisWorkbook.Worksheets("ATND PASTE HERE")
    Set shSc2Gbb = ThisWorkbook.Worksheets("SC2GBB")
    Set shFront = ThisWorkbook.Worksheets("Front")
    jmlBrAtnd = atnd.Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 6 To jmlBrAtnd
'        abisIpId
        atnd.Cells(i, 4).Copy
        shSc2Gbb.Cells(i - 4, 8).PasteSpecial xlPasteValues
'        vlanId
        atnd.Cells(i, 12).Copy
        shSc2Gbb.Cells(i - 4, 11).PasteSpecial xlPasteValues
'        ipAbis
        atnd.Cells(i, 13).Copy
        shSc2Gbb.Cells(i - 4, 12).PasteSpecial xlPasteValues
'        ipGateway
        atnd.Cells(i, 14).Copy
        shSc2Gbb.Cells(i - 4, 13).PasteSpecial xlPasteValues
'        NetworkPrefixLength
        atnd.Cells(i, 15).Copy
        shSc2Gbb.Cells(i - 4, 14).PasteSpecial xlPasteValues
'        IPTimeServer1
        atnd.Cells(i, 16).Copy
        shSc2Gbb.Cells(i - 4, 15).PasteSpecial xlPasteValues
'        IPTimeServer2
        atnd.Cells(i, 17).Copy
        shSc2Gbb.Cells(i - 4, 16).PasteSpecial xlPasteValues
    Next i
    Application.CutCopyMode = False
    
End Sub
Sub cpFrn()
    Dim jmlsc2gbb As Integer
    Dim jmlFrn As Integer
    Dim shSc2Gbb As Worksheet
    Dim shFront As Worksheet
    Dim i As Integer
    Set shSc2Gbb = ThisWorkbook.Worksheets("SC2GBB")
    Set shFront = ThisWorkbook.Worksheets("Front")
    
    jmlsc2gbb = shSc2Gbb.Range("A" & Rows.Count).End(xlUp).Row
    jmlFrn = shFront.Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 13 To jmlFrn
'        SiteName
        shFront.Range("B2").Copy
        shSc2Gbb.Cells(i - 11, 1).PasteSpecial xlPasteValues
        
'        GsmSector
        shFront.Cells(i, 1).Copy
        shSc2Gbb.Cells(i - 11, 2).PasteSpecial xlPasteValues
        
'        Trx
        If shFront.Cells(i, 4).Value = "8" Then
            shSc2Gbb.Cells(i - 11, 4).Value = "2"
        
        ElseIf shFront.Cells(i, 4).Value = "0" Then
            shSc2Gbb.Cells(i - 11, 4).Value = "1"
        Else
            shSc2Gbb.Cells(i - 11, 4).Value = ""
        End If
        
'        frequencyBand
        shFront.Cells(i, 15).Copy
        shSc2Gbb.Cells(i - 11, 5).PasteSpecial xlPasteValues
        
'        userLabel
        shFront.Range("B9").Copy
        shSc2Gbb.Cells(i - 11, 6).PasteSpecial xlPasteValues

    Next i
    shSc2Gbb.Range("A2:P1000").HorizontalAlignment = xlCenter
    shSc2Gbb.Columns.autoFit
    shSc2Gbb.Rows.autoFit
    Application.CutCopyMode = False
End Sub
Sub CrtScr2GinBB(control As IRibbonControl)
'Sub CrtScr2GinBB()
    TurnOnSpeed True
    Dim jmlabisIpId As Long
    Dim ShSc2GinBB As Worksheet
    
    Set ShSc2GinBB = ThisWorkbook.Worksheets("SC2GBB")
    path = "C:"
    
    jmlabisIpId = ShSc2GinBB.Range("H" & Rows.Count).End(xlUp).Row
    jmlBrsector = ShSc2GinBB.Range("B" & Rows.Count).End(xlUp).Row
'    jmlBrCovRel = Worksheets("CovRel").Range("E" & Rows.Count).End(xlUp).Row
'    jmlBrU2U = Worksheets("utran_rel").Range("H" & Rows.Count).End(xlUp).Row
'    jmlBrgsmrel = Worksheets("gsmrel").Range("B" & Rows.Count).End(xlUp).Row
'
'    XIublink = 2
'    Xutranrel = 2
    Xncell = 2
'    Xgsmrel = 2
'    XCovRel = 2
    
    For Cntatnd = 2 To jmlabisIpId
        Debug.Print Cntatnd
        Debug.Print jmlabisIpId
        
        sname = ShSc2GinBB.Cells(Cntatnd, 8).Value
        freqb = ShSc2GinBB.Cells(Cntatnd, 5).Value
        Debug.Print freqb & sname

        Set fs = CreateObject("Scripting.FileSystemObject")
        If fs.FolderExists(path & "\Script") = False Then
            fs.CreateFolder (path & "\Script")
        End If
        
        If fs.FolderExists(path & "\Script\" & sname) = False Then
            fs.CreateFolder (path & "\Script\" & sname)
        End If
        
        D = Format(Date, "yyyymmdd") & "_Script_2GInBB"
    
        fileName = path & "\Script\" & sname

        Call Scr2gInBB
        
    Next Cntatnd
'    Worksheets("3G CDR PASTE HERE").Select
    TurnOnSpeed False
    MsgBox "Completed Generating Script" & Chr(10) & "The Path Script in " & path & "\Script\"
End Sub
Sub Testing()
Dim ShSc2GinBB As Worksheet
Dim jmlBrAtnd As Integer, jmlBrGsmSec As Integer
Dim idTrx As Integer, sEFunctionR18 As Integer, sEFunctionR9 As Integer
Dim arfcnMax As String, arfcnMin As String
Dim fBand As String

Set ShSc2GinBB = ThisWorkbook.Worksheets("SC2GBB")
jmlBrAtnd = ShSc2GinBB.Range("H" & Rows.Count).End(xlUp).Row
jmlBrGsmSec = ShSc2GinBB.Range("A" & Rows.Count).End(xlUp).Row
Dim sEFunctionRBuf(2 To 500) As Integer

'For Cntatnd = 2 To jmlBrAtnd
Print #1, "//================================================================"
Print #1, "// Script Add 2G in BB & Trx                        "
Print #1, "// +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
Print #1, "// SITE ID        :  " & ShSc2GinBB.Cells(Cntatnd, 8).Value
Print #1, "// Generated on :  " & Format(Date, "dddd, yyyy-mm-dd")
Print #1, "// Generated by :  ABC_Tools by TAC XL Team                  "
Print #1, "//================================================================"
Print #1, Chr(10)

Print #1, "//================================================================"
Print #1, "// Transport                                                      "
Print #1, "//================================================================"

    idTrx = 0
'    Transport Abis
    Print #1, "gs+"
    Print #1, "crn Transport=1,VlanPort=Abis"
    Print #1, "egressQosClassification"
    Print #1, "egressQosMarking"
    Print #1, "egressQosQueueMap"
    Print #1, "encapsulation EthernetPort=" & ShSc2GinBB.Cells(Cntatnd, 9).Value
    Print #1, "ingressQosMarking"
    Print #1, "isTagged true"
    Print #1, "lowLatencySwitching false"
    Print #1, "userLabel VLAN_ABIS"
    Print #1, "vlanId " & ShSc2GinBB.Cells(Cntatnd, 11).Value
    Print #1, "end"
    Print #1, ""
'    Router Abis
    Print #1, "crn Transport=1,Router=vr_Abis"
    Print #1, "hopLimit 64"
    Print #1, "pathMtuExpiresIPv6 86400"
    Print #1, "routingPolicyLocal"
    Print #1, "ttl 64"
    Print #1, "userLabel"
    Print #1, "end"
    Print #1, ""
'    InterfaceIPv4=1
    Print #1, "crn Transport=1,Router=vr_Abis,InterfaceIPv4=1"
    Print #1, "aclEgress"
    Print #1, "aclIngress"
    Print #1, "arpTimeout 300"
    Print #1, "bfdProfile"
    Print #1, "bfdStaticRoutes 0"
    Print #1, "egressQosMarking"
    Print #1, "encapsulation VlanPort=Abis"
    Print #1, "ingressQosMarking"
    Print #1, "loopback false"
    Print #1, "mtu 1500"
    Print #1, "pcpArp 6"
    Print #1, "routesHoldDownTimer"
    Print #1, "routingPolicyIngress"
    Print #1, "userLabel ROUTER_ABIS"
    Print #1, "end"
    Print #1, ""
'    AddressIPv4=1
    Print #1, "crn Transport=1,Router=vr_Abis,InterfaceIPv4=1,AddressIPv4=1"
    Print #1, "address " & ShSc2GinBB.Cells(Cntatnd, 12).Value & "/" & _
    ShSc2GinBB.Cells(Cntatnd, 14)
    Print #1, "configurationMode 0"
    Print #1, "dhcpClientIdentifier"
    Print #1, "dhcpClientIdentifierType 0"
    Print #1, "userLabel"
    Print #1, "end"
    Print #1, ""
'    RouteTableIPv4Static=1
    Print #1, "crn Transport=1,Router=vr_Abis,RouteTableIPv4Static=1"
    Print #1, "end"
    Print #1, ""
    
'    RouteTableIPv4Static=1,Dst=1
    Print #1, "crn Transport=1,Router=vr_Abis,RouteTableIPv4Static=1,Dst=1"
    Print #1, "dst 0.0.0.0/0"
    Print #1, "end"
    Print #1, ""
    
'    RouteTableIPv4Static=1,Dst=1,NextHop=1
    Print #1, "crn Transport=1,Router=vr_Abis,RouteTableIPv4Static=1,Dst=1,NextHop=1"
    Print #1, "address " & ShSc2GinBB.Cells(Cntatnd, 13).Value
    Print #1, "adminDistance 1"
    Print #1, "bfdMonitoring true"
    Print #1, "discard false"
    Print #1, "reference"
    Print #1, "end"
    Print #1, ""
    
'    Ntp=1,NtpFrequencySync=5
    Print #1, "crn Transport=1,Ntp=1,NtpFrequencySync=5"
    Print #1, "addressIPv4Reference Router=vr_Abis,InterfaceIPv4=1,AddressIPv4=1"
    Print #1, "dscp 54"
    Print #1, "syncServerNtpIpAddress " & ShSc2GinBB.Cells(Cntatnd, 15).Value
    Print #1, "end"
    Print #1, ""
    
'    Ntp=1,NtpFrequencySync=6
    Print #1, "crn Transport=1,Ntp=1,NtpFrequencySync=6"
    Print #1, "addressIPv4Reference Router=vr_Abis,InterfaceIPv4=1,AddressIPv4=1"
    Print #1, "dscp 54"
    Print #1, "syncServerNtpIpAddress " & ShSc2GinBB.Cells(Cntatnd, 16).Value
    Print #1, "end"
    Print #1, ""
    
'    RadioEquipmentClockReference=5
    Print #1, "crn Transport=1,Synchronization=1,RadioEquipmentClock=1," & _
    "RadioEquipmentClockReference=5"
    Print #1, "adminQualityLevel qualityLevelValueOptionI=2,qualityLevelValueOptionII=2," & _
    "qualityLevelValueOptionIII=1"
    Print #1, "administrativeState 1"
    Print #1, "encapsulation Transport=1,Ntp=1,NtpFrequencySync=5"
    Print #1, "holdOffTime 1000"
    Print #1, "priority 5"
    Print #1, "useQLFrom 1"
    Print #1, "waitToRestoreTime 60"
    Print #1, "end"
    Print #1, ""
    
'    RadioEquipmentClockReference=6
    Print #1, "crn Transport=1,Synchronization=1,RadioEquipmentClock=1," & _
    "RadioEquipmentClockReference=6"
    Print #1, "adminQualityLevel qualityLevelValueOptionI=2,qualityLevelValueOptionII=2," & _
    "qualityLevelValueOptionIII=1"
    Print #1, "administrativeState 1"
    Print #1, "encapsulation Transport=1,Ntp=1,NtpFrequencySync=6"
    Print #1, "holdOffTime 1000"
    Print #1, "priority 6"
    Print #1, "useQLFrom 1"
    Print #1, "waitToRestoreTime 60"
    Print #1, "end"
    Print #1, "gs-"
    Print #1, ""
    
    Print #1, "//================================================================"
    Print #1, "// BtsFunction=1                                                  "
    Print #1, "//================================================================"
'    BtsFunction=1
    Print #1, "gs+"
    Print #1, "crn BtsFunction=1"
'    if ShSc2GinBB.Cells(Cntatnd, 8).Value
    For CntGsmSec = 2 To jmlBrGsmSec
        If ShSc2GinBB.Cells(Cntatnd, 8).Value = ShSc2GinBB.Cells(CntGsmSec, 1).Value Then
        Dim uselblBtsF As String
            userlblBtsF = ShSc2GinBB.Cells(CntGsmSec, 6).Value
        End If
    Next CntGsmSec
    Print #1, "userLabel " & userlblBtsF
    Print #1, "end"
    Print #1, "gs-"
    Print #1, ""
    
    idTrx = 0
    Dim plusSEF As Integer, parX As Integer
    Dim sctr As Integer
    plusSEF = 2
    parX = 1
    
    
    sctr = 1
    For CntGsmSec = 2 To jmlBrGsmSec

        
        If ShSc2GinBB.Cells(Cntatnd, 8).Value = ShSc2GinBB.Cells(CntGsmSec, 1).Value Then
            sEFunctionRBuf(plusSEF) = parX
        Print #1, "//================================================================"
        Print #1, "// Sector " & (sctr)
        Print #1, "//================================================================"
            sctr = sctr + 1
'           ##SECTOR##
'           ###GsmSector=JK46170   ##Ganti##
            Print #1, "crn BtsFunction=1,GsmSector=" & ShSc2GinBB.Cells(CntGsmSec, 2)
            Print #1, "userLabel " & userlblBtsF
            Print #1, "end"
            Print #1, ""
'           ###GsmSector=JK46170,AbisIp=1
            Print #1, "crn BtsFunction=1,GsmSector=" & ShSc2GinBB.Cells(CntGsmSec, 2) & _
            ",AbisIp=1"
            Print #1, "administrativeState 0"
            Print #1, "bscBrokerIpAddress " & ShSc2GinBB.Cells(Cntatnd, 10).Value
            Print #1, "dscpSectorControlUL 46"
            Print #1, "gsmSectorName " & ShSc2GinBB.Cells(CntGsmSec, 2)
            Print #1, "initialRetransmissionPeriod 1"
            Print #1, "ipv4Address Router=vr_Abis,InterfaceIPv4=1,AddressIPv4=1"
            Print #1, "keepAlivePeriod 1"
            Print #1, "maxRetransmission 5"
            Print #1, "retransmissionCap 4"
            Print #1, "userLabel "
            Print #1, "end"
            Print #1, ""
            
'           ###BtsFunction=1,GsmSector=JK46170,Trx=0
            Dim TrxSTG As String
            Dim i As Integer
            TrxSTG = ShSc2GinBB.Cells(CntGsmSec, 4).Value
            If TrxSTG = "1" Then
                For i = 1 To 1
                    Print #1, "crn BtsFunction=1,GsmSector=" & ShSc2GinBB.Cells(CntGsmSec, 2) & "," & _
                    "Trx=" & idTrx
                    Print #1, "administrativeState 0"
            
                    If ShSc2GinBB.Cells(CntGsmSec, 5).Value = "GSM1800" Then
                        arfcnMax = "622"
                        arfcnMin = "512"
                        fBand = "3"
                        sEFunctionR = 3 + sEFunctionRBuf(plusSEF)
                    ElseIf ShSc2GinBB.Cells(CntGsmSec, 5).Value = "GSM900" Then
                        arfcnMax = "99"
                        arfcnMin = "91"
                        fBand = "8"
                        sEFunctionR = 0 + sEFunctionRBuf(plusSEF)
                    End If
            
                    Print #1, "arfcnMax " & arfcnMax
                    Print #1, "arfcnMin " & arfcnMin
                    Print #1, "combinedCellType 0"
                    Print #1, "configuredMaxTxPower 10000"
                    Print #1, "frequencyBand " & fBand
                    Print #1, "noOfRxAntennas 2"
                    Print #1, "noOfTxAntennas 2"
                    Print #1, "reservedMaxTxPower"
                    Print #1, "rfBranchRxRef"
                    Print #1, "rfBranchTxRef"
                    Print #1, "rxImbAlarmThreshold 60"
                    Print #1, "rxImbMinNoOfSamples 8010"
                    Print #1, "rxImbSupWindowSize 288"
                    Print #1, "sectorEquipmentFunctionRef SectorEquipmentFunction=" & sEFunctionR
                    Print #1, "userLabel"
                    Print #1, "end"
                    Print #1, ""
                    idTrx = idTrx + 1
                
                Next i
            ElseIf TrxSTG = "2" Then
                For i = 1 To 2
                    Print #1, "crn BtsFunction=1,GsmSector=" & ShSc2GinBB.Cells(CntGsmSec, 2) & "," & _
                    "Trx=" & idTrx
                    Print #1, "administrativeState 0"
            
                    If ShSc2GinBB.Cells(CntGsmSec, 5).Value = "GSM1800" Then
                        arfcnMax = "622"
                        arfcnMin = "512"
                        fBand = "3"
                        sEFunctionR = 3 + sEFunctionRBuf(plusSEF)
                    ElseIf ShSc2GinBB.Cells(CntGsmSec, 5).Value = "GSM900" Then
                        arfcnMax = "99"
                        arfcnMin = "91"
                        fBand = "8"
                        sEFunctionR = 0 + sEFunctionRBuf(plusSEF)
                    End If
            
                    Print #1, "arfcnMax " & arfcnMax
                    Print #1, "arfcnMin " & arfcnMin
                    Print #1, "combinedCellType 0"
                    Print #1, "configuredMaxTxPower 10000"
                    Print #1, "frequencyBand " & fBand
                    Print #1, "noOfRxAntennas 2"
                    Print #1, "noOfTxAntennas 2"
                    Print #1, "reservedMaxTxPower"
                    Print #1, "rfBranchRxRef"
                    Print #1, "rfBranchTxRef"
                    Print #1, "rxImbAlarmThreshold 60"
                    Print #1, "rxImbMinNoOfSamples 8010"
                    Print #1, "rxImbSupWindowSize 288"
                    Print #1, "sectorEquipmentFunctionRef SectorEquipmentFunction=" & sEFunctionR
                    Print #1, "userLabel"
                    Print #1, "end"
                    Print #1, ""
                    idTrx = idTrx + 1
                
               Next i
            
            
            End If
            
            idTrx = 0

            parX = parX + 1
'        Else
'                MsgBox "Please makesure SiteName and abisIpId match"
'                Application.Cursor = xlDefault
'                End
'                Application.Cursor = xlDefault
            
        End If
    
    Next CntGsmSec
    Print #1, "//================================================================"
    Print #1, "// Feature                                                        "
    Print #1, "//================================================================"
    Print #1, "ldeb BtsFunction=1$"
    Print #1, "set CXC4012017|CXC4012026|CXC4012037|CXC4012021 FeatureState 1"
    Print #1, ""
    Print #1, "//================================================================"
    Print #1, "// CV                                                             "
    Print #1, "//================================================================"
    Print #1, "$date = `date +%y%m%d_%H%M%S`"
    Print #1, "cvms $nodename_$date_Add2G"
    Print #1, ""
'Next Cntatnd
End Sub
Sub Scr2gInBB()
Dim ShSc2GinBB As Worksheet
Dim jmlBrAtnd As Integer, jmlBrGsmSec As Integer

Set ShSc2GinBB = ThisWorkbook.Worksheets("SC2GBB")
jmlBrAtnd = ShSc2GinBB.Range("H" & Rows.Count).End(xlUp).Row
jmlBrGsmSec = ShSc2GinBB.Range("A" & Rows.Count).End(xlUp).Row

fileNamedoc = fileName & "\" & sname & "_" & freqb & "_" & D & ".mos"
Debug.Print fileNamedoc
Open fileNamedoc For Output As #1
Close #1: Kill (fileNamedoc)
Open fileNamedoc For Output As #1
Dim alline As String
 
Sheets("SC2GBB").Select



Call Testing
 
Print #1, alline
Close #1
End Sub

Attribute VB_Name = "Module1_Update"
'================================================================
' Swerdlow Cockpit Model - UPDATE_Phase + versioning (FIFO cap)
' Auto-generated field map: 391 flattened inputs.
' Database layout: row 4 = headers, data from row 5.
' Meta cols: A=RecordID B=PhaseID C=PhaseName D=Timestamp
'            E=UpdatedBy F=Notes ; first input field = column 7 (G).
'================================================================
Option Explicit

Private Const HDR_ROW As Long = 4
Private Const FIRST_DATA As Long = 5
Private Const FIRST_FIELD_COL As Long = 7
Private Const PHASENAME_ADDR As String = "$B$13"

' Returns the Cockpit input-cell addresses in Database-column order.
Public Function FieldAddrs() As Variant
    Dim A(391) As String
    A(1) = "$B$13"
    A(2) = "$B$14"
    A(3) = "$B$15"
    A(4) = "$B$16"
    A(5) = "$B$17"
    A(6) = "$B$18"
    A(7) = "$B$19"
    A(8) = "$B$20"
    A(9) = "$B$21"
    A(10) = "$B$24"
    A(11) = "$B$25"
    A(12) = "$B$26"
    A(13) = "$B$27"
    A(14) = "$B$28"
    A(15) = "$B$29"
    A(16) = "$B$30"
    A(17) = "$B$32"
    A(18) = "$B$33"
    A(19) = "$A$37"
    A(20) = "$B$37"
    A(21) = "$C$37"
    A(22) = "$D$37"
    A(23) = "$E$37"
    A(24) = "$A$38"
    A(25) = "$B$38"
    A(26) = "$C$38"
    A(27) = "$D$38"
    A(28) = "$E$38"
    A(29) = "$A$39"
    A(30) = "$B$39"
    A(31) = "$C$39"
    A(32) = "$D$39"
    A(33) = "$E$39"
    A(34) = "$A$40"
    A(35) = "$B$40"
    A(36) = "$C$40"
    A(37) = "$D$40"
    A(38) = "$E$40"
    A(39) = "$A$41"
    A(40) = "$B$41"
    A(41) = "$C$41"
    A(42) = "$D$41"
    A(43) = "$E$41"
    A(44) = "$A$42"
    A(45) = "$B$42"
    A(46) = "$C$42"
    A(47) = "$D$42"
    A(48) = "$E$42"
    A(49) = "$A$43"
    A(50) = "$B$43"
    A(51) = "$C$43"
    A(52) = "$D$43"
    A(53) = "$E$43"
    A(54) = "$A$44"
    A(55) = "$B$44"
    A(56) = "$C$44"
    A(57) = "$D$44"
    A(58) = "$E$44"
    A(59) = "$A$45"
    A(60) = "$B$45"
    A(61) = "$C$45"
    A(62) = "$D$45"
    A(63) = "$E$45"
    A(64) = "$A$46"
    A(65) = "$B$46"
    A(66) = "$C$46"
    A(67) = "$D$46"
    A(68) = "$E$46"
    A(69) = "$A$47"
    A(70) = "$B$47"
    A(71) = "$C$47"
    A(72) = "$D$47"
    A(73) = "$E$47"
    A(74) = "$A$48"
    A(75) = "$B$48"
    A(76) = "$C$48"
    A(77) = "$D$48"
    A(78) = "$E$48"
    A(79) = "$A$49"
    A(80) = "$B$49"
    A(81) = "$C$49"
    A(82) = "$D$49"
    A(83) = "$E$49"
    A(84) = "$A$50"
    A(85) = "$B$50"
    A(86) = "$C$50"
    A(87) = "$D$50"
    A(88) = "$E$50"
    A(89) = "$A$51"
    A(90) = "$B$51"
    A(91) = "$C$51"
    A(92) = "$D$51"
    A(93) = "$E$51"
    A(94) = "$B$56"
    A(95) = "$B$58"
    A(96) = "$B$59"
    A(97) = "$B$60"
    A(98) = "$B$61"
    A(99) = "$B$62"
    A(100) = "$B$63"
    A(101) = "$B$64"
    A(102) = "$B$65"
    A(103) = "$B$66"
    A(104) = "$B$70"
    A(105) = "$C$70"
    A(106) = "$D$70"
    A(107) = "$B$71"
    A(108) = "$C$71"
    A(109) = "$D$71"
    A(110) = "$B$72"
    A(111) = "$C$72"
    A(112) = "$D$72"
    A(113) = "$B$73"
    A(114) = "$C$73"
    A(115) = "$D$73"
    A(116) = "$B$76"
    A(117) = "$B$78"
    A(118) = "$C$78"
    A(119) = "$B$79"
    A(120) = "$C$79"
    A(121) = "$B$80"
    A(122) = "$C$80"
    A(123) = "$B$81"
    A(124) = "$C$81"
    A(125) = "$B$82"
    A(126) = "$C$82"
    A(127) = "$B$83"
    A(128) = "$C$83"
    A(129) = "$B$85"
    A(130) = "$C$85"
    A(131) = "$B$86"
    A(132) = "$C$86"
    A(133) = "$B$87"
    A(134) = "$C$87"
    A(135) = "$C$88"
    A(136) = "$C$89"
    A(137) = "$B$93"
    A(138) = "$B$94"
    A(139) = "$B$98"
    A(140) = "$B$99"
    A(141) = "$B$100"
    A(142) = "$B$101"
    A(143) = "$B$102"
    A(144) = "$B$103"
    A(145) = "$A$106"
    A(146) = "$B$106"
    A(147) = "$D$106"
    A(148) = "$E$106"
    A(149) = "$F$106"
    A(150) = "$H$106"
    A(151) = "$I$106"
    A(152) = "$J$106"
    A(153) = "$K$106"
    A(154) = "$L$106"
    A(155) = "$N$106"
    A(156) = "$O$106"
    A(157) = "$Q$106"
    A(158) = "$A$107"
    A(159) = "$B$107"
    A(160) = "$D$107"
    A(161) = "$E$107"
    A(162) = "$F$107"
    A(163) = "$H$107"
    A(164) = "$I$107"
    A(165) = "$J$107"
    A(166) = "$K$107"
    A(167) = "$L$107"
    A(168) = "$N$107"
    A(169) = "$O$107"
    A(170) = "$Q$107"
    A(171) = "$A$108"
    A(172) = "$B$108"
    A(173) = "$D$108"
    A(174) = "$E$108"
    A(175) = "$F$108"
    A(176) = "$H$108"
    A(177) = "$I$108"
    A(178) = "$J$108"
    A(179) = "$K$108"
    A(180) = "$L$108"
    A(181) = "$N$108"
    A(182) = "$O$108"
    A(183) = "$Q$108"
    A(184) = "$A$109"
    A(185) = "$B$109"
    A(186) = "$D$109"
    A(187) = "$E$109"
    A(188) = "$F$109"
    A(189) = "$H$109"
    A(190) = "$I$109"
    A(191) = "$J$109"
    A(192) = "$K$109"
    A(193) = "$L$109"
    A(194) = "$N$109"
    A(195) = "$O$109"
    A(196) = "$Q$109"
    A(197) = "$A$110"
    A(198) = "$B$110"
    A(199) = "$D$110"
    A(200) = "$E$110"
    A(201) = "$F$110"
    A(202) = "$H$110"
    A(203) = "$I$110"
    A(204) = "$J$110"
    A(205) = "$K$110"
    A(206) = "$L$110"
    A(207) = "$N$110"
    A(208) = "$O$110"
    A(209) = "$Q$110"
    A(210) = "$A$111"
    A(211) = "$B$111"
    A(212) = "$D$111"
    A(213) = "$E$111"
    A(214) = "$F$111"
    A(215) = "$H$111"
    A(216) = "$I$111"
    A(217) = "$J$111"
    A(218) = "$K$111"
    A(219) = "$L$111"
    A(220) = "$N$111"
    A(221) = "$O$111"
    A(222) = "$Q$111"
    A(223) = "$A$112"
    A(224) = "$B$112"
    A(225) = "$D$112"
    A(226) = "$E$112"
    A(227) = "$F$112"
    A(228) = "$H$112"
    A(229) = "$I$112"
    A(230) = "$J$112"
    A(231) = "$K$112"
    A(232) = "$L$112"
    A(233) = "$N$112"
    A(234) = "$O$112"
    A(235) = "$Q$112"
    A(236) = "$A$113"
    A(237) = "$B$113"
    A(238) = "$D$113"
    A(239) = "$E$113"
    A(240) = "$F$113"
    A(241) = "$H$113"
    A(242) = "$I$113"
    A(243) = "$J$113"
    A(244) = "$K$113"
    A(245) = "$L$113"
    A(246) = "$N$113"
    A(247) = "$O$113"
    A(248) = "$Q$113"
    A(249) = "$A$114"
    A(250) = "$B$114"
    A(251) = "$D$114"
    A(252) = "$E$114"
    A(253) = "$F$114"
    A(254) = "$H$114"
    A(255) = "$I$114"
    A(256) = "$J$114"
    A(257) = "$K$114"
    A(258) = "$L$114"
    A(259) = "$N$114"
    A(260) = "$O$114"
    A(261) = "$Q$114"
    A(262) = "$A$115"
    A(263) = "$B$115"
    A(264) = "$D$115"
    A(265) = "$E$115"
    A(266) = "$F$115"
    A(267) = "$H$115"
    A(268) = "$I$115"
    A(269) = "$J$115"
    A(270) = "$K$115"
    A(271) = "$L$115"
    A(272) = "$N$115"
    A(273) = "$O$115"
    A(274) = "$Q$115"
    A(275) = "$A$116"
    A(276) = "$B$116"
    A(277) = "$D$116"
    A(278) = "$E$116"
    A(279) = "$F$116"
    A(280) = "$H$116"
    A(281) = "$I$116"
    A(282) = "$J$116"
    A(283) = "$K$116"
    A(284) = "$L$116"
    A(285) = "$N$116"
    A(286) = "$O$116"
    A(287) = "$Q$116"
    A(288) = "$A$117"
    A(289) = "$B$117"
    A(290) = "$D$117"
    A(291) = "$E$117"
    A(292) = "$F$117"
    A(293) = "$H$117"
    A(294) = "$I$117"
    A(295) = "$J$117"
    A(296) = "$K$117"
    A(297) = "$L$117"
    A(298) = "$N$117"
    A(299) = "$O$117"
    A(300) = "$Q$117"
    A(301) = "$A$118"
    A(302) = "$B$118"
    A(303) = "$D$118"
    A(304) = "$E$118"
    A(305) = "$F$118"
    A(306) = "$H$118"
    A(307) = "$I$118"
    A(308) = "$J$118"
    A(309) = "$K$118"
    A(310) = "$L$118"
    A(311) = "$N$118"
    A(312) = "$O$118"
    A(313) = "$Q$118"
    A(314) = "$A$119"
    A(315) = "$B$119"
    A(316) = "$D$119"
    A(317) = "$E$119"
    A(318) = "$F$119"
    A(319) = "$H$119"
    A(320) = "$I$119"
    A(321) = "$J$119"
    A(322) = "$K$119"
    A(323) = "$L$119"
    A(324) = "$N$119"
    A(325) = "$O$119"
    A(326) = "$Q$119"
    A(327) = "$A$120"
    A(328) = "$B$120"
    A(329) = "$D$120"
    A(330) = "$E$120"
    A(331) = "$F$120"
    A(332) = "$H$120"
    A(333) = "$I$120"
    A(334) = "$J$120"
    A(335) = "$K$120"
    A(336) = "$L$120"
    A(337) = "$N$120"
    A(338) = "$O$120"
    A(339) = "$Q$120"
    A(340) = "$B$122"
    A(341) = "$I$127"
    A(342) = "$J$127"
    A(343) = "$K$127"
    A(344) = "$L$127"
    A(345) = "$M$127"
    A(346) = "$I$128"
    A(347) = "$J$128"
    A(348) = "$K$128"
    A(349) = "$L$128"
    A(350) = "$M$128"
    A(351) = "$I$129"
    A(352) = "$J$129"
    A(353) = "$K$129"
    A(354) = "$L$129"
    A(355) = "$M$129"
    A(356) = "$I$130"
    A(357) = "$J$130"
    A(358) = "$K$130"
    A(359) = "$L$130"
    A(360) = "$M$130"
    A(361) = "$I$131"
    A(362) = "$J$131"
    A(363) = "$K$131"
    A(364) = "$L$131"
    A(365) = "$M$131"
    A(366) = "$B$136"
    A(367) = "$E$136"
    A(368) = "$B$137"
    A(369) = "$E$137"
    A(370) = "$B$138"
    A(371) = "$E$138"
    A(372) = "$B$139"
    A(373) = "$E$139"
    A(374) = "$B$140"
    A(375) = "$E$140"
    A(376) = "$B$141"
    A(377) = "$E$141"
    A(378) = "$B$142"
    A(379) = "$B$144"
    A(380) = "$E$144"
    A(381) = "$B$145"
    A(382) = "$E$145"
    A(383) = "$B$146"
    A(384) = "$E$146"
    A(385) = "$B$147"
    A(386) = "$E$147"
    A(387) = "$B$148"
    A(388) = "$E$148"
    A(389) = "$B$149"
    A(390) = "$E$149"
    A(391) = "$B$150"
    Dim out() As String, i As Long
    ReDim out(1 To 391)
    For i = 1 To 391
        out(i) = A(i)
    Next i
    FieldAddrs = out
End Function

Public Sub UPDATE_Phase()
    Dim wsDB As Worksheet, wsCk As Worksheet
    Set wsDB = ThisWorkbook.Sheets("Database")
    Set wsCk = ThisWorkbook.Sheets("Cockpit")

    Dim phase As String
    phase = CStr(wsCk.Range("SelectedPhase").Value)
    If Len(Trim(phase)) = 0 Then
        MsgBox "Select a Phase before updating.", vbExclamation: Exit Sub
    End If

    Dim lastRow As Long
    lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    If lastRow < FIRST_DATA Then lastRow = HDR_ROW
    Dim newRow As Long: newRow = lastRow + 1

    Dim newID As Long
    If lastRow >= FIRST_DATA Then
        newID = Application.WorksheetFunction.Max(wsDB.Range("A" & FIRST_DATA & ":A" & lastRow)) + 1
    Else
        newID = 1
    End If

    Application.ScreenUpdating = False
    wsDB.Cells(newRow, 1).Value = newID
    wsDB.Cells(newRow, 2).Value = phase
    wsDB.Cells(newRow, 3).Value = wsCk.Range(PHASENAME_ADDR).Value
    wsDB.Cells(newRow, 4).Value = Now
    wsDB.Cells(newRow, 5).Value = Environ("Username")
    wsDB.Cells(newRow, 6).Value = wsCk.Range("Cockpit_Notes").Value

    Dim addrs As Variant, i As Long
    addrs = FieldAddrs()
    For i = LBound(addrs) To UBound(addrs)
        wsDB.Cells(newRow, FIRST_FIELD_COL + (i - LBound(addrs))).Value = _
            wsCk.Range(addrs(i)).Value
    Next i

    EnforceCap wsDB, phase
    Application.ScreenUpdating = True
    MsgBox "Saved record #" & newID & " for phase " & phase & ".", vbInformation
End Sub

' FIFO versioning cap: keep at most MaxRecordsPerPhase records per Phase;
' overflow rows are archived (recoverable) then removed.
Private Sub EnforceCap(wsDB As Worksheet, phase As String)
    Dim cap As Long
    cap = CLng(wsDB.Range("MaxRecordsPerPhase").Value)
    If cap < 1 Then cap = 1

    Do
        Dim lastRow As Long
        lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
        If lastRow < FIRST_DATA Then Exit Sub

        Dim cnt As Long
        cnt = Application.WorksheetFunction.CountIf( _
              wsDB.Range("B" & FIRST_DATA & ":B" & lastRow), phase)
        If cnt <= cap Then Exit Do

        Dim r As Long, minID As Double, minRow As Long
        minID = 1E+18: minRow = 0
        For r = FIRST_DATA To lastRow
            If CStr(wsDB.Cells(r, 2).Value) = phase Then
                If wsDB.Cells(r, 1).Value < minID Then
                    minID = wsDB.Cells(r, 1).Value: minRow = r
                End If
            End If
        Next r
        If minRow = 0 Then Exit Do
        ArchiveRow wsDB, minRow
        wsDB.Rows(minRow).Delete
    Loop
End Sub

Private Sub ArchiveRow(wsDB As Worksheet, srcRow As Long)
    Dim wsA As Worksheet
    On Error Resume Next
    Set wsA = ThisWorkbook.Sheets("Archive")
    On Error GoTo 0
    If wsA Is Nothing Then Exit Sub          ' archive disabled -> hard delete only
    Dim destRow As Long
    destRow = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row + 1
    If destRow < 4 Then destRow = 4
    wsDB.Rows(srcRow).Copy
    wsA.Rows(destRow).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
End Sub

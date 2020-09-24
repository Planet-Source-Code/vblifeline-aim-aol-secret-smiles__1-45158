Attribute VB_Name = "Module1"
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public SmileCode(102)
Public SmileCode2(15)
Public SmileStart(102)
Public tCat

Public Function IsCompiled() As Boolean

    On Error GoTo Result
    Debug.Print (1 / 0) 'this won't work if compiled
    IsCompiled = True
    Exit Function
Result:
    IsCompiled = False

End Function


Public Function LoadCode(nTemp)

    ncode = "<font sml=" & Chr(34) & SmileCode(tCat) & Chr(34) & ">" & SmileCode2(nTemp) & "</font>"
    Form1.Text1.Text = Form1.Text1.Text & ncode

End Function


Public Function LoadStart()

    Form2.Picture1.Height = Form2.Image1.Height
    Form2.Picture1.Width = Form2.Image1.Width
    Form2.Picture1.Picture = Form2.Image1.Picture

    For X = 0 To 102
        BitBlt Form1.pIcon(X).hDC, 0, 0, 19, 19, Form2.Picture1.hDC, 0 * 19 + 3, SmileStart(X), &HCC0020
        Form1.pIcon(X).Refresh
    Next


End Function


Public Function LoadCat(mCat)

    tCat = mCat

    For X = 0 To 15
        BitBlt Form1.pIcon2(X).hDC, 0, 0, 19, 19, Form2.Picture1.hDC, X * 19 + 3, SmileStart(mCat), &HCC0020
        Form1.pIcon2(X).Refresh
    Next


End Function


Public Function SmileStartIt()

    SmileCode2(0) = ":-D"
    SmileCode2(1) = "=-O"
    SmileCode2(2) = ":-*"
    SmileCode2(3) = ">:o"
    SmileCode2(4) = "8-)"
    SmileCode2(5) = ":-$"
    SmileCode2(6) = ":-!"
    SmileCode2(7) = ":-["
    SmileCode2(8) = "O:-)"
    SmileCode2(9) = ":'("
    SmileCode2(10) = ":-X"
    SmileCode2(11) = ":)"
    SmileCode2(12) = ":("
    SmileCode2(13) = ";-)"
    SmileCode2(14) = ":-P"
    SmileCode2(15) = ":-\"
    SmileCode(0) = "AgHSBfc="
    SmileCode(1) = "AgHSBjw="
    SmileCode(2) = "AgHSBxo="
    SmileCode(3) = "AgHSBxI="
    SmileCode(4) = "AgHSBxZ="
    SmileCode(5) = "AgHSBdQ="
    SmileCode(6) = "AgHSBdh="
    SmileCode(7) = "AgHSBdo="
    SmileCode(8) = "AgHSBx4="
    SmileCode(9) = "AgHSByI="
    SmileCode(10) = "AgHSByY="
    SmileCode(11) = "AgHSBcg="
    SmileCode(12) = "AgHSBdA="
    SmileCode(13) = "AgHSD9A="
    SmileCode(14) = "AgHSENQ="
    SmileCode(15) = "AgHSED8="
    SmileCode(16) = "AgHSEQ8="
    SmileCode(17) = "AgHSEtI="
    SmileCode(18) = "AgHSEMw="
    SmileCode(19) = "AgHSETY="
    SmileCode(20) = "AgHSE6U="
    SmileCode(21) = "AgHSD90="
    SmileCode(22) = "AgHSEWU="
    SmileCode(23) = "AgHSETI="
    SmileCode(24) = "AgHSEMI="
    SmileCode(25) = "AgHSD80="
    SmileCode(26) = "AgHSE5E="
    SmileCode(27) = "AgHSE6M="
    SmileCode(28) = "AgHSD9o="
    SmileCode(29) = "AgHSFIA="
    SmileCode(30) = "AgHSEJM="
    SmileCode(31) = "AgHSEJ0="
    SmileCode(32) = "AgHSEKA="
    SmileCode(33) = "AgHSEMU="
    SmileCode(34) = "AgHSE1g="
    SmileCode(35) = "AgHSFGs="
    SmileCode(36) = "AgHSD9U="
    SmileCode(37) = "AgHSEtg="
    SmileCode(38) = "AgHSEU0="
    SmileCode(39) = "AgHSE5I="
    SmileCode(40) = "AgHSE6Q="
    SmileCode(41) = "AgHSD8o="
    SmileCode(42) = "AgHSFGY="
    SmileCode(43) = "AgHSEGY="
    SmileCode(44) = "AgHSEDo="
    SmileCode(45) = "AgHSEsw="
    SmileCode(46) = "AgHSEV8="
    SmileCode(47) = "AgHSEtA="
    SmileCode(48) = "AgHSFKo="
    SmileCode(49) = "AgHSD6U="
    SmileCode(50) = "AgHSD7k="
    SmileCode(51) = "AgHSEto="
    SmileCode(52) = "AgHSEc0="
    SmileCode(53) = "AgHSD4Y="
    SmileCode(54) = "AgHSEcU="
    SmileCode(55) = "AgHSHoo="
    SmileCode(56) = "AgHSHok="
    SmileCode(57) = "AgHSGU8="
    SmileCode(58) = "AgHSHCc="
    SmileCode(59) = "AgHSHCY="
    SmileCode(60) = "AgHSHuk="
    SmileCode(61) = "AgHSF5k="
    SmileCode(62) = "AgHSG/s="
    SmileCode(63) = "AgHSHo8="
    SmileCode(64) = "AgHSGUQ="
    SmileCode(65) = "AgHSGUE="
    SmileCode(66) = "AgHSGT0="
    SmileCode(67) = "AgHSF9U="
    SmileCode(68) = "AgHSHpE="
    SmileCode(69) = "AgHSGTc="
    SmileCode(70) = "AgHSHv8="
    SmileCode(71) = "AgHSHuo="
    SmileCode(72) = "AgHSHo4="
    SmileCode(73) = "AgHSFjo="
    SmileCode(74) = "AgHSHv4="
    SmileCode(75) = "AgHSF5g="
    SmileCode(76) = "AgHSHpA="
    SmileCode(77) = "AgHSHo0="
    SmileCode(78) = "AgHSGWU="
    SmileCode(79) = "AgHSF54="
    SmileCode(80) = "AgHSF5s="
    SmileCode(81) = "AgHSGRM="
    SmileCode(82) = "AgHSF8c="
    SmileCode(83) = "AgHSGWI="
    SmileCode(84) = "AgHSF/U="
    SmileCode(85) = "AgHSGSk="
    SmileCode(86) = "AgHSGSY="
    SmileCode(87) = "AgHSGR4="
    SmileCode(88) = "AgHSGds="
    SmileCode(89) = "AgHSHuY="
    SmileCode(90) = "AgHSHpI="
    SmileCode(91) = "AgHSF5Y="
    SmileCode(92) = "AgHSF5M="
    SmileCode(93) = "AgHSHus="
    SmileCode(94) = "AgHSH2s="
    SmileCode(95) = "AgHSH2U="
    SmileCode(96) = "AgHSH2o="
    SmileCode(97) = "AgHSH1A="
    SmileCode(98) = "AgHSJmk="
    SmileCode(99) = "AgHSJFc="
    SmileCode(100) = "AgHSJmw="
    SmileCode(101) = "AgHSJnE="
    SmileCode(102) = "AgHSJnQ="
    SmileStart(0) = "4"
    SmileStart(1) = "23"
    SmileStart(2) = "42"
    SmileStart(3) = "61"
    SmileStart(4) = "80"
    SmileStart(5) = "99"
    SmileStart(6) = "118"
    SmileStart(7) = "137"
    SmileStart(8) = "156"
    SmileStart(9) = "175"
    SmileStart(10) = "194"
    SmileStart(11) = "213"
    SmileStart(12) = "232"
    SmileStart(13) = "253"
    SmileStart(14) = "273"
    SmileStart(15) = "293"
    SmileStart(16) = "313"
    SmileStart(17) = "335"
    SmileStart(18) = "357"
    SmileStart(19) = "378"
    SmileStart(20) = "402"
    SmileStart(21) = "419"
    SmileStart(22) = "441"
    SmileStart(23) = "462"
    SmileStart(24) = "483"
    SmileStart(25) = "504"
    SmileStart(26) = "525"
    SmileStart(27) = "546"
    SmileStart(28) = "567"
    SmileStart(29) = "588"
    SmileStart(30) = "610"
    SmileStart(31) = "631"
    SmileStart(32) = "652"
    SmileStart(33) = "673"
    SmileStart(34) = "695"
    SmileStart(35) = "716"
    SmileStart(36) = "736"
    SmileStart(37) = "758"
    SmileStart(38) = "780"
    SmileStart(39) = "802"
    SmileStart(40) = "823"
    SmileStart(41) = "845"
    SmileStart(42) = "867"
    SmileStart(43) = "889"
    SmileStart(44) = "912"
    SmileStart(45) = "934"
    SmileStart(46) = "956"
    SmileStart(47) = "978"
    SmileStart(48) = "1002"
    SmileStart(49) = "1024"
    SmileStart(50) = "1046"
    SmileStart(51) = "1068"
    SmileStart(52) = "1090"
    SmileStart(53) = "1112"
    SmileStart(54) = "1134"
    SmileStart(55) = "1156"
    SmileStart(56) = "1177"
    SmileStart(57) = "1198"
    SmileStart(58) = "1220"
    SmileStart(59) = "1241"
    SmileStart(60) = "1262"
    SmileStart(61) = "1283"
    SmileStart(62) = "1304"
    SmileStart(63) = "1325"
    SmileStart(64) = "1345"
    SmileStart(65) = "1365"
    SmileStart(66) = "1386"
    SmileStart(67) = "1407"
    SmileStart(68) = "1429"
    SmileStart(69) = "1451"
    SmileStart(70) = "1472"
    SmileStart(71) = "1494"
    SmileStart(72) = "1515"
    SmileStart(73) = "1536"
    SmileStart(74) = "1557"
    SmileStart(75) = "1578"
    SmileStart(76) = "1599"
    SmileStart(77) = "1620"
    SmileStart(78) = "1642"
    SmileStart(79) = "1663"
    SmileStart(80) = "1684"
    SmileStart(81) = "1705"
    SmileStart(82) = "1726"
    SmileStart(83) = "1747"
    SmileStart(84) = "1768"
    SmileStart(85) = "1788"
    SmileStart(86) = "1809"
    SmileStart(87) = "1829"
    SmileStart(88) = "1850"
    SmileStart(89) = "1871"
    SmileStart(90) = "1892"
    SmileStart(91) = "1913"
    SmileStart(92) = "1934"
    SmileStart(93) = "1955"
    SmileStart(94) = "1976"
    SmileStart(95) = "1997"
    SmileStart(96) = "2020"
    SmileStart(97) = "2040"
    SmileStart(98) = "2062"
    SmileStart(99) = "2083"
    SmileStart(100) = "2105"
    SmileStart(101) = "2127"
    SmileStart(102) = "2149"

End Function




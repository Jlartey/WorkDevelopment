
Dim msg
msg = ""
processcode

Sub processcode()
Dim tbl, kf, rec, vld, st, aAmt, rAmt, aAmt2, usr, cDt, rDt, hrs, hrMax, hrMax2, isSup, cHrMax
tbl = UCase(GetPageVariable("TableName"))
cDt = Now()
'hrMax = 24 * 3
' hrMax = 24 * 7
hrMax = 10000 '' 24 * 3
hrMax2 = 2400
rec = Trim(Request("inpReceiptID"))
vld = False
If Len(Trim(rec)) > 0 Then
  aAmt = Trim(Request("inpTransProcessValue1"))
  rAmt = Trim(Request("inpTransProcessValue2"))
  usr = Trim(GetComboNameFld("Receipt", rec, "SystemUserID"))
  If Len(usr) > 0 Then
    isSup = IsCashierSupervisor(jSchd)
    cHrMax = hrMax
    If isSup Then
      cHrMax = hrMax2
    End If
    If (UCase(usr) = UCase(Trim(uName)) Or isSup) Then
      rDt = GetComboNameFld("Receipt", rec, "ReceiptDate")
      hrs = DateDiff("h", CDate(rDt), cDt)
      If (hrs <= hrMax) Or (isSup And (hrs <= hrMax2)) Then
        If IsNumeric(rAmt) And CDbl(rAmt) > 0 Then  ' @ Peter - 23 Oct 2023 - Added validation for refund amount > 0
           aAmt2 = GetComboNameFld("Receipt", rec, "ReceiptAmount3")
          If IsNumeric(aAmt) And IsNumeric(rAmt) And IsNumeric(aAmt2) Then
            If Round(CDbl(aAmt), 2) = Round(CDbl(aAmt2), 2) Then
              If CDbl(rAmt) <= CDbl(aAmt) Then
                vld = True
              Else
                msg = "Refund Amount [" & rAmt & "] is more than Available Refund Amount [" & aAmt & "]."
              End If
            Else
              msg = "Available Refund Amount [" & aAmt & "] is different from what is in the system [" & aAmt2 & "]."
            End If
          Else
            msg = "Refund Amount [" & rAmt & "] must be a number."
          End If
        Else
          msg = "Refund Amount must be greater than 0."
        End If
      Else
        msg = "You cannot Refund this Receipt [" & UCase(rec) & "] because it was issued more than [" & CStr(cHrMax) & "] hours [" & CStr(hrs) & "] ago > [" & FormatDateDetail(rDt) & "]"
      End If

    Else
      msg = "You are not the Cashier [" & GetComboName("SystemUser", usr) & "] who issued this Receipt [" & UCase(rec) & "]"
    End If
  Else
    msg = "The Cashier name for this receipt [" & UCase(rec) & "] is not VALID"
  End If
Else
  msg = "Receipt No. is blank."
End If

  ' ''If UCase(jSchd)=UCase(uName) Then
  ' If UCase(jSchd)=UCase("SystemAdmin") Or UCase(jSchd)=UCase(uName) Or UCase(GetComboNameFld("SystemUser", uName, "StaffID"))=UCase("STF001") Then
  '   vld = False
  '   SetPageMessages "Policy Alert! Admin Disabled"
  ' End If

 If Not vld Then
  If objPage.rtnHdlProcessPoint Then
    objPage.hdlProcessPoint = False
    SetPageMessages msg
  End If
 End If
End Sub

Function IsCashierSupervisor(jb)
  Dim arr, ul, num, lst, ot
  ot = False
  lst = "ChiefCashier||CreditControl||M11"
  arr = Split(lst, "||")
  ul = UBound(arr)
  For num = 0 To ul
    If UCase(Trim(arr(num))) = UCase(Trim(jb)) Then
      ot = True
      Exit For
    End If
  Next
  IsCashierSupervisor = ot
End Function





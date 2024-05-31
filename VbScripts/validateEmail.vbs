' This is a business rule

SetPageMessages "First Business Rule"
EmailValidation

Sub EmailValidation()
    Dim emailInput, graGov, match, vld
    vld = True
    
    emailInput = Trim(Request("InpClearanceMemo2"))
    graGov = "@gra.gov.gh"
    
    Set regex = New regExp
    regex.Pattern = graGov
    
    Set match = regex.Execute(emailInput)
    
    If match.count = 0 Then
        SetPageMessages "Enter a valid email"
        vld = False
    Else
        SetPageMessages "You have entered a valid email"
    End If
    If Not vld Then
        If objPage.rtnHdlProcessPoint Then
            objPage.hdlProcessPoint = False
        End If
    End If
End Sub
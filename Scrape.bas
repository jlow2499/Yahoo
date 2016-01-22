Sub YahooFinance()

    Const cURL = "http://finance.yahoo.com/"
    
    Dim IE As InternetExplorer
    Dim doc As HTMLDocument
    Dim cap As HTMLInputElement
    Dim LoginForm As HTMLFormElement
    Dim HTMLelement As IHTMLElement
    Dim SYMBOL As HTMLInputElement
    Dim CAPRES As HTMLInputElement
    Dim SYMBOLEXCEL As Object
    Dim irow As Integer
    Set IE = New InternetExplorer
    LastRow = ActiveSheet.UsedRange.Rows.Count
    
    IE.Visible = True
    
    IE.navigate cURL
       
    Do While IE.ReadyState <> READYSTATE_COMPLETE Or IE.Busy: DoEvents: Loop
    
Do
    
    irow = irow + 1
    
    If Range("A" & irow).Value = "" Then
    Application.StatusBar = "SCRUB COMPLETE"
    MsgBox "RUN COMPLETE"
    Exit Sub
    End If
    
    Set SYMBOLEXCEL = Range("A" & irow)
    Set SYMBOL = IE.Document.getelementbyid("txtQuotes")
    SYMBOL.Value = SYMBOLEXCEL
    
    Application.StatusBar = "Processing " & irow & " of " & LastRow
    
    IE.Document.getelementbyid("btnQuotes").Click
    
    'Application.Wait (Now + TimeValue("0:00:05"))
    
    Do While IE.ReadyState <> READYSTATE_COMPLETE Or IE.Busy: DoEvents: Loop

On Error Resume Next
    
    Application.Wait (Now + TimeValue("0:00:02"))
    
    Range("B" & irow) = IE.Document.getelementbyid("table1").getelementsbytagname("td")(0).innerText
    Range("C" & irow) = IE.Document.getelementbyid("table1").getelementsbytagname("td")(1).innerText
    Range("D" & irow) = IE.Document.getelementbyid("table1").getelementsbytagname("td")(2).innerText
    Range("E" & irow) = IE.Document.getelementbyid("table1").getelementsbytagname("td")(3).innerText
    Range("F" & irow) = IE.Document.getelementbyid("table2").getelementsbytagname("td")(0).innerText
    Range("G" & irow) = IE.Document.getelementbyid("table2").getelementsbytagname("td")(1).innerText
    Range("H" & irow) = IE.Document.getelementbyid("table2").getelementsbytagname("td")(2).innerText
    Range("I" & irow) = IE.Document.getelementbyid("table2").getelementsbytagname("td")(3).innerText

    Loop
   
End Sub


#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

Pause::Pause
^F12::

;********************************************
; USER VARIABLES
; Please Change the variables below
;********************************************
userName := "jwiens10"
userPassword := "Winter18"
userWorkbook := "C:\Users\jason\github\oi-puller-test.xlsx"
userSaveAsLocation := "C:\Users\jason\github\Open-Invoice-Puller\images\"

; Application Variables
counter := 2

; Create an instance of IE and Open
wb := ComObjCreate("InternetExplorer.Application")
wb.Visible := True

; Navigate to OI
loginURL := "https://www.openinvoice.com/docp/public/OILogin.xhtml"
wb.Navigate(loginURL)
While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for site to load
    sleep, 100

; Enter username and password and submit
wb.document.getElementByID("j_username").Value := userName
wb.document.getElementByID("j_password").Value := userPassword
sleep, 100
wb.document.getElementByID("loginBtn").click()
While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for site to load
    sleep, 100	

; Open workbook
wbk := ComObjGet(userWorkbook)

; loop through all invoices listed in workbook
Loop {
    statusMsg := ""

    ; get invoice url and text (usually a doc numebr or some sort of unique id) from excel
    invoiceURL := wbk.Sheets("Sheet1").Cells(counter, 1).Value
    invoiceText := wbk.Sheets("Sheet1").Cells(counter, 2).Value

    ; if invoiceURL is empty then terminate loop
    if(!invoiceURL) {
        break
    }

    ; navigate to invoice
    wb.Navigate(invoiceURL)
    While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for site to load
        sleep, 100

    ; check invoice exists
    if(!wb.document.getElementByID("pageTitle")) {
        wbk.Sheets("Sheet1").Cells(counter, 3).Value := "Does not exist or Permission Denied"
        counter += 1
        continue
    }

    ; get a list of all attachments
    links := wb.document.getElementByID("DIV_JOURNAL_attachments").getElementsByTagName("A")

    ; download and save electronic invoice
    send, ^p
    WinWaitActive,Print
    send, {enter}
    WinWaitActive,Save PDF File As
    Sleep, 300
    send, %userSaveAsLocation%%invoiceText%{enter}
    WinWaitActive, ahk_class AcrobatSDIWindow
    send, ^q
    WinWaitClose, ahk_class AcrobatSDIWindow
    statusMsg .= "Invoice"

    ; if attachements exist download and save each attachment
    if (links.Length > 0) {
        Loop % links.Length {
            links[A_Index-1].focus()
            Sleep, 200
            Send, +{F10}
            Sleep, 300
            Send, a
            WinWaitActive,Save As
            Sleep, 500
            name := links[A_Index-1].innerText
            send, %userSaveAsLocation%%invoiceText% - %name%{enter}
            WinWaitClose,Save As
        }
        statusMsg .= " and Attachements"
    }

    ; print status to workbook
    wbk.Sheets("Sheet1").Cells(counter, 3).Value := statusMsg . " Successfully Downloaded"

    ; increment counter
    counter += 1
}

; close ie, send finish message and end app
wb.quit()
MsgBox, Finished
return





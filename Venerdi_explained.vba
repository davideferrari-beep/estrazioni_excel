' Subroutine that automates SAP job extraction for "Venerdì" sheet.
Sub Venerdì()
    Dim SapGuiAuto As Object ' Holds reference to SAP GUI automation entry point.
    Dim sapApp     As Object ' Represents the SAP GUI scripting engine application.
    Dim Connection As Object ' Represents the current SAP connection.
    Dim Session    As Object ' Represents the active SAP GUI session.
    Dim data, job, path, name, usr, trx, prg, flnm As String ' Variables storing row values from Excel.
    Dim i As Long ' Row index for iterating over worksheet entries.
    Dim ws As Worksheet ' Worksheet reference for the "VEN" sheet.
    
    
' Imposta il foglio di lavoro - Sets the worksheet context for processing.
    Set ws = ThisWorkbook.Sheets("VEN") ' Acquire the worksheet named "VEN".

' Inizia dalla seconda riga - Start iterating from row 2 to skip headers.
    i = 2 ' Start reading data from the second row of the sheet.

' Continua il loop finché la cella in colonna A non è vuota - Loop until column A is empty.
    Do While ws.Cells(i, 1).Value <> "" ' Iterate while column A contains a job name.
        Application.DisplayAlerts = False ' Disable Excel alert popups during automation.
        Application.AlertBeforeOverwriting = False ' Prevent overwrite confirmation dialogs.
' Assicurati che questa sia la data corretta. Se deve variare, usa ws.Cells(i, 1).Value - Note about adjusting date source.
        data = ws.Range("A1").Value ' Retrieve the reference date stored in cell A1.
        job = ws.Cells(i, 1).Value ' Fetch the job name from column A of the current row.
        path = ws.Cells(i, 2).Value ' Fetch the download path from column B.
        name = ws.Cells(i, 3).Value ' Fetch the file name from column C.
        usr = ws.Cells(i, 4).Value ' Fetch the SAP user from column D.
        trx = ws.Cells(i, 5).Value ' Fetch the SAP transaction code from column E.
        prg = ws.Cells(i, 6).Value ' Fetch the program name from column F.
        flnm = ws.Cells(i, 7).Value ' Fetch the file pattern filter from column G.
' Inizializza la sessione SAP - Prepare SAP GUI automation objects.
        Set SapGuiAuto = GetObject("SAPGUI") ' Access the running SAP GUI automation object.
        Set sapApp = SapGuiAuto.GetScriptingEngine ' Obtain the scripting engine instance.
        Set Connection = sapApp.Children(0) ' Use the first available SAP connection.
        Set Session = Connection.Children(0) ' Use the first session within that connection.

' Porta SAP in uno stato iniziale richiamando la transazione SM37 - Reset SAP context before commands.
        Session.findById("wnd[0]").maximize ' Maximize the main SAP window for clarity.
        Session.findById("wnd[0]/tbar[0]/okcd").Text = trx ' Insert the transaction code into the command field.
        Session.findById("wnd[0]").sendVKey 0 ' Execute the transaction by pressing Enter.
        
        If InStr(1, trx, "SM", vbTextCompare) > 0 Then ' Check if the transaction is an SM-type (job monitoring) transaction.
    ' Inserisce il valore della data e del job presi da Excel - Populate SM37 filters.
            Session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").Text = job ' Set the job name filter.
            Session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").Text = usr ' Set the user filter.
            Session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").Text = data ' Set the from-date filter.
            Session.findById("wnd[0]/usr/ctxtBTCH2170-TO_DATE").Text = data ' Set the to-date filter.
            Session.findById("wnd[0]/usr/ctxtBTCH2170-TO_DATE").SetFocus ' Focus the to-date field to ensure proper input.
            Session.findById("wnd[0]/usr/ctxtBTCH2170-TO_DATE").caretPosition = 6 ' Move the caret to the end of the date.
            Session.findById("wnd[0]").sendVKey 0 ' Confirm the selection with Enter.
            Session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").caretPosition = 23 ' Position the caret at the end of the job name.
            Session.findById("wnd[0]").sendVKey 8 ' Execute the selection criteria (F8).
            Session.findById("wnd[0]/usr/chk[1,13]").Selected = True ' Select the "Spool" checkbox in the job list.
            Session.findById("wnd[0]/usr/lbl[37,13]").SetFocus ' Focus the related label for navigation.
            Session.findById("wnd[0]/usr/lbl[37,13]").caretPosition = 0 ' Reset caret position in the label.
            Session.findById("wnd[0]").sendVKey 2 ' Open the job log via Shift+F2 (or equivalent command).
            Session.findById("wnd[0]/usr/chk[1,3]").Selected = True ' Select the desired spool request line.
            Session.findById("wnd[0]/usr/lbl[14,3]").SetFocus ' Focus the label to prepare for action.
            Session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0 ' Reset caret on the label.
            Session.findById("wnd[0]").sendVKey 2 ' Trigger the display/download options.
            Session.findById("wnd[0]").sendVKey 48 ' Invoke the download/save function (Shift+F12).
            If InStr(1, name, "pos", vbTextCompare) > 0 Then ' If file name suggests a positional format, choose alternate output.
                    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select ' Select the "Text with Tabs" option.
                    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus ' Ensure the radio button remains selected.
                Else ' Otherwise default to standard format.
                    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").Select ' Select the "Spreadsheet" or primary option.
            End If
            Session.findById("wnd[1]").sendVKey 0 ' Confirm the output format selection.
            Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = path ' Set the download directory path.
            Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = name ' Set the target file name.
            Session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = Len(name) ' Position caret at end of file name for confirmation.
            Session.findById("wnd[1]/tbar[0]/btn[11]").press ' Press the save/confirm button to download.
            Session.findById("wnd[0]").sendVKey 12 ' Navigate back to the previous screen (Shift+F3).
            Session.findById("wnd[0]").sendVKey 12 ' Repeat navigation back to reach start screen.
            Session.findById("wnd[0]").sendVKey 12 ' Continue closing intermediate screens.
            Session.findById("wnd[0]").sendVKey 12 ' Ensure return to main screen ready for next iteration.
            Else ' Handle transactions that are not SM-type.
                Session.findById("wnd[0]/usr/txtP_LOCAL").Text = name ' Provide local file name parameter.
                Session.findById("wnd[0]/usr/ctxtP_FILE").Text = prg ' Set the program name parameter.
                Session.findById("wnd[0]/usr/txtP_PATERN").Text = flnm ' Set the file name pattern to filter files.
                Session.findById("wnd[0]/usr/txtP_PATERN").SetFocus ' Focus the pattern field for validation.
                Session.findById("wnd[0]/usr/txtP_PATERN").caretPosition = 7 ' Place caret at end of pattern string.
                Session.findById("wnd[0]").sendVKey 8 ' Execute the selection or run the program.
                Session.findById("wnd[0]/usr/chk[0,2]").Selected = True ' Select the desired data row in the results list.
                Session.findById("wnd[0]/tbar[1]/btn[5]").press ' Press the button to start download/export.
                Session.findById("wnd[1]/tbar[0]/btn[0]").press ' Confirm save in the popup dialog.
                Session.findById("wnd[0]").sendVKey 3 ' Exit the transaction (F3).
                Session.findById("wnd[0]").sendVKey 12 ' Return to the SAP main screen.

        End If
        ' Passa alla riga successiva - Advance to next job configuration.
        i = i + 1 ' Move to the next worksheet row for processing.
    Loop
    Application.DisplayAlerts = True ' Re-enable Excel alerts after automation completes.

End Sub ' End of the Venerdì automation routine.


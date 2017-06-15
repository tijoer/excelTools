Sub Mail_Workbook()
 ' From https://accountingmacros.com

 Dim OutApp As Object
 Dim OutMail As Object
 
 Set OutApp = CreateObject("Outlook.Application")
 Set OutMail = OutApp.CreateItem(0)
 
 On Error Resume Next
 With OutMail
 .To = ""
 .CC = ""
 .BCC = ""
 .Subject = ActiveWorkbook.Name
 .body = "See attached."
 .Attachments.Add ActiveWorkbook.FullName
 .Display
 End With
 On Error GoTo 0
 
 Set OutMail = Nothing
 Set OutApp = Nothing
End Sub

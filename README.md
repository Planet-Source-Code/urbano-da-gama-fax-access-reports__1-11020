<div align="center">

## Fax Access Reports


</div>

### Description

Use this code to fax enable your application. Use this code to fax your access report. It also explains how to use winfax. I hope you find this code useful and spend a moment to rate it.
 
### More Info
 
You will need to have symantec Winfax Pro10.0 installed on your machine before executing this code.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[urbano da gama](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/urbano-da-gama.md)
**Level**          |Intermediate
**User Rating**    |4.6 (41 globes from 9 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/urbano-da-gama-fax-access-reports__1-11020/archive/master.zip)





### Source Code

```
Public Function FaxReport() As Boolean
  On Error GoTo EH
  Dim lReport As Report
  Dim lFileName As String
  Dim lSendObj As Object' winfax send object
  Dim lRet As Long
  'delete any existing fax report file
  lFileName = CurDir & "\" & "FaxReport.html"
  If Dir(lFileName) <> vbNullString Then
    Kill lFileName
  End If
  'save as an html file so that it can be faxed
  'as an attachement
  DoCmd.OutputTo acOutputReport, _
      mReportName, "html", lFileName
  'now is the time to fax the html file
  Set lSendObj = CreateObject("WinFax.SDKSend")
  lRet = lSendObj.SetAreaCode("801")
  lRet = lSendObj.SetCountryCode("1")
  lRet = lSendObj.SetNumber(9816661)
  lRet = lSendObj.AddRecipient()
  lRet = lSendObj.AddAttachmentFile(lFileName)
  lRet = lSendObj.ShowCallProgress(1)
  lRet = lSendObj.Send(0)
  lRet = lSendObj.Done()
  Exit Function
EH:
  Exit Function
end function
```


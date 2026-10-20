Attribute VB_Name = "ExportForWebTool"
' Saves the three sheets the Climate Risk TARP Finder reads into a "data" folder
' next to this workbook: one CSV per sheet, plus erp-data.js, which the web page loads.
' Run it (Developer > Macros > ExportSheetsForWebTool) after editing the workbook.
Sub ExportSheetsForWebTool()
    Dim sheetsToExport As Variant, fileNames As Variant, i As Integer
    Dim folder As String, js As String, csvText As String
    sheetsToExport = Array("Database", "TARP & ERP", "Project Questionaire")
    fileNames = Array("Database.csv", "TARP_ERP.csv", "Project_Questionaire.csv")
    folder = ThisWorkbook.Path & Application.PathSeparator & "data"
    If Dir(folder, vbDirectory) = "" Then MkDir folder
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    js = "/* Generated from the ERP workbook by the ExportForWebTool macro. Do not edit by hand. */" & vbLf & "window.ERP_CSV = {" & vbLf
    For i = LBound(sheetsToExport) To UBound(sheetsToExport)
        ThisWorkbook.Sheets(sheetsToExport(i)).Copy
        ActiveWorkbook.SaveAs Filename:=folder & Application.PathSeparator & fileNames(i), FileFormat:=xlCSVUTF8
        ActiveWorkbook.Close SaveChanges:=False
        csvText = ReadUtf8(folder & Application.PathSeparator & fileNames(i))
        js = js & """" & sheetsToExport(i) & """: `" & JsEscape(csvText) & "`"
        If i < UBound(sheetsToExport) Then js = js & ","
        js = js & vbLf
    Next i
    js = js & "};" & vbLf
    WriteUtf8 folder & Application.PathSeparator & "erp-data.js", js
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Exported the Database, TARP & ERP and Project Questionaire sheets to " & folder, vbInformation
End Sub

Private Function JsEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, "`", "\`")
    s = Replace(s, "${", "\${")
    JsEscape = s
End Function

Private Function ReadUtf8(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText
    st.Close
End Function

Private Sub WriteUtf8(ByVal path As String, ByVal text As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText text
    st.SaveToFile path, 2
    st.Close
End Sub

Option Explicit

LaunchMacro

Sub LaunchMacro()
    Dim xl
    Dim xlBook
    Dim sCurPath

    On Error Resume Next

    sCurPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")
    Set xl = CreateObject("Excel.Application")
     Set xlBook = xl.Workbooks.Open(sCurPath & "\master_v0.60.xlsm", 0, True)
      xl.Application.Visible = True
      xl.Application.Run "master_v0.60.xlsm!import_results_module.importResults"
      xl.DisplayAlerts = False
      xlBook.Saved = True
      xlBook.Close

      Set xlBook = Nothing
      Set xl = Nothing
End Sub

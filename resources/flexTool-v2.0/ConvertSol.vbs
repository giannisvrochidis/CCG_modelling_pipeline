Option Explicit

LaunchMacro

Sub LaunchMacro()
    Dim xl
    Dim xlBook
    Dim sCurPath
    Dim args
    Dim fn
    Dim folder
    Set args = Wscript.Arguments
    folder = args(0)
    fn = args(1)
    sCurPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")
    Set xl = CreateObject("Excel.Application")
    Set xlBook = xl.Workbooks.Open(sCurPath & "\master_v0.60.xlsm", 0, True)
    xl.Application.Visible = True
    xl.Application.Run "master_v0.60.xlsm!functions.convert_sol", CStr(folder), CStr(fn)
    xl.DisplayAlerts = False
    xlBook.Saved = True
    xlBook.Close
    xl.Quit
      Set xlBook = Nothing
      Set xl = Nothing
End Sub

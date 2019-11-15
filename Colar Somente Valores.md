Essa função faz com seja possivel somente colar em valores

```
Option Explicit

 Private Sub Worksheet_Change(ByVal Target As Range)
    Dim UndoList As String
 
    Application.ScreenUpdating = False
    Application.EnableEvents = False
 
    On Error GoTo Whoa
    
    UndoList = Application.CommandBars("Standard").Controls("&Desfazer").List(1)
  
    If Left(UndoList, 5) <> "Colar" And UndoList <> "Preenchimento Automático" _
    Then GoTo LetsContinue
 
    Application.Undo
 
    If UndoList = "Preenchimento Automático" Then Selection.Copy

    On Error Resume Next
 
    Target.Select
    ActiveSheet.PasteSpecial Format:="Texto", _
    Link:=False, DisplayAsIcon:=False
 
    Target.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
    On Error GoTo 0
    
    Union(Target, Selection).Select
 
LetsContinue:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
Whoa:
    MsgBox Err.Description
    Resume LetsContinue
End Sub

``` vbs

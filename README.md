<div align="center">

## WinKill


</div>

### Description

WinKill destroys a window if you know its title bar caption.
 
### More Info
 
' Create a form a text box called txtName and a command button called cmdKill


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Grove](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-grove.md)
**Level**          |Unknown
**User Rating**    |4.1 (65 globes from 16 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-grove-winkill__1-886/archive/master.zip)

### API Declarations

```
' Api constants for General DeclarationsConst WM_DESTROY = &H2
Const WM_CLOSE = &H10' Api Functions for general declarations
Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Private Declare Function SendMessageA Lib "user32" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
```


### Source Code

```
'*************************************************************************
'WinKill Form Code
'*************************************************************************
Private Function Kill(hWnd&)
 Dim Res& ' Ask it politely to close
 Res = SendMessageA(hWnd, WM_CLOSE, 0, 0)
 ' Kill it (just in case)
 Res = SendMessageA(hWnd, WM_DESTROY, 0, 0)
End Function
Private Sub cmdKill_Click()
 Dim hWnd& ' Get the window handle
 hWnd = FindWindowA(vbNullString, txtName.Text) ' Call the kill function
 Kill (hWnd)
End Sub
```


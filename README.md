<div align="center">

## Prevent more than one instance from any directory


</div>

### Description

After seeing a few examples submitted attempting to prevent more than one instance. I decided to submit my version.

This code will only prevent the 2nd instance from running no matter what directory both .exe are executed from.

In addition, it does not use App.PrevInstance and will not bring the 1st instance to the front, just merely a *WORKING* example of how to stop more then one instance from running and the code is very easy to understand and use in your own apps.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sabro](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sabro.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sabro-prevent-more-than-one-instance-from-any-directory__1-65747/archive/master.zip)

### API Declarations

```
Private Const ERROR_ALREADY_EXISTS As Long = 183&amp;
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As Any, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
```


### Source Code

```
Private Function IsPrevInstance() As Boolean
  Dim lngMutex As Long
  lngMutex = CreateMutex(ByVal 0&, 1, App.Title)
  If (Err.LastDllError = ERROR_ALREADY_EXISTS) Then
    ReleaseMutex lngMutex
    CloseHandle lngMutex
    IsPrevInstance = True
  Else
    IsPrevInstance = False
  End If
End Function
Private Sub Form_Load()
  If IsPrevInstance = True Then
    MsgBox "This Program Is Already Active!"
    End
  End If
End Sub
```


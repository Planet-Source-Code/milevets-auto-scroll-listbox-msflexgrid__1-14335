<div align="center">

## Auto Scroll ListBox/MSFlexGrid


</div>

### Description

Ever wanted to automatically scroll (vertical) ListBox or MSFlexGrid? Below is a quick solution and it's especially handy when you cannot use object.TopRow (due to variable row height, etc).

Happy Coding
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[milevets](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/milevets.md)
**Level**          |Intermediate
**User Rating**    |4.9 (34 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/milevets-auto-scroll-listbox-msflexgrid__1-14335/archive/master.zip)

### API Declarations

```
'Import Win32 API function
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Constant
Const WM_VSCROLL = &H115
Const SB_BOTTOM = 7
```


### Source Code

```
'Call below where auto scroll is intended
SendMessage MSFG.hwnd, WM_VSCROLL, SB_BOTTOM, 0
'MSFG is my FlexGrid control, can be changed to ListBox
```


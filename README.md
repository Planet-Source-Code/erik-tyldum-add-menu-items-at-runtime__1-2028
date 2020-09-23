<div align="center">

## Add menu items at runtime\! :\)


</div>

### Description

This code adds allows you to add menu item's on your form while you are running it(runtime :)...
 
### More Info
 
'Goto the menu Editor!

'Add a menu, and a submenu.

'Name the submenu: 'mnuTest' and set its index property to '0' <- Important!

'Then you click OK!:)

'Add a Command Button, named Command1... and paste my code into it, and


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Erik Tyldum](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/erik-tyldum.md)
**Level**          |Unknown
**User Rating**    |4.0 (24 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/erik-tyldum-add-menu-items-at-runtime__1-2028/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
    Load mnuTest(mnuTest.Count)
End Sub
```


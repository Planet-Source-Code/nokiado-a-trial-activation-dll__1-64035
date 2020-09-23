<div align="center">

## A Trial activation DLL


</div>

### Description

This is a secure trial activation system which gives users of your program 30 days to try out your program before they can buy using there unique computer ID. This software also comes with Clock-Back detection. Please leave comments and vote. Thanks.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2006-01-14 13:02:26
**By**             |[NokiaDO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/nokiado.md)
**Level**          |Intermediate
**User Rating**    |3.8 (23 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[A\_Trial\_ac1965151142006\.zip](https://github.com/Planet-Source-Code/nokiado-a-trial-activation-dll__1-64035/archive/master.zip)

### API Declarations

```
Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not TrialActivation.EstaRegistrado Then
 If TrialActivation.PrevInstancia Then Exit Sub
 TrialActivation.Show
End If
If TrialActivation.Finalizar Then End
End Sub
```






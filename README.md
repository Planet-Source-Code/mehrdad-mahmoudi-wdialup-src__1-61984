<div align="center">

## wDialup SRC


</div>

### Description

Its a RAS Entry recoverer that returns all your RAS entries in windows NT platform (XP/2000/2003/Vista and ...)  also returns: VPN, Dialup, ADSL and ... Entries !!!

It uses LSA Secret APIs to recover RAS Entries

Its an class and you can use it easily in your code ...

but don't forget copyright ... :D

Special thanks to SNE (website: http://vbnet.ru) !!!

Regards,

w0lf
 
### More Info
 
NT RAS Entries


<span>             |<span>
---                |---
**Submitted On**   |2005-07-30 07:20:04
**By**             |[Mehrdad Mahmoudi](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mehrdad-mahmoudi.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[wDialup\_SR1919057302005\.zip](https://github.com/Planet-Source-Code/mehrdad-mahmoudi-wdialup-src__1-61984/archive/master.zip)

### API Declarations

```
Private Declare Function LsaRetrievePrivateData Lib "ADVAPI32.dll" (ByVal PolicyHandle As Long, ByRef KeyName As LSA_UNICODE_STRING, ByVal PrivateData As Long) As Long
```






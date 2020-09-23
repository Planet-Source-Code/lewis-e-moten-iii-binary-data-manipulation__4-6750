<div align="center">

## Binary Data Manipulation


</div>

### Description

These two functions allow you to convert between Unicode and Ascii strings. This is great if you are working with the Request.BinaryRead/BinaryWrite methods or binary data within a database.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lewis E\. Moten III](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lewis-e-moten-iii.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Files](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files__4-2.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lewis-e-moten-iii-binary-data-manipulation__4-6750/archive/master.zip)





### Source Code

```
Function AsciiToUnicode(ByRef pstrAscii)
	Dim llngLength
	Dim llngIndex
	Dim llngAscii
	Dim lstrUnicode
	llngLength = LenB(pstrAscii)
	For llngIndex = 1 to llngLength
		llngAscii = AscB(MidB(pstrAscii, llngIndex, 1))
		lstrUnicode = lstrUnicode & Chr(llngAscii)
	Next
	AsciiToUnicode = lstrUnicode
End Function
Function UnicodeToAscii(ByRef pstrUnicode)
	Dim llngLength
	Dim llngIndex
	Dim llngAscii
	Dim lstrAscii
	llngLength = Len(pstrUnicode)
	For llngIndex = 1 to llngLength
		llngAscii = Asc(Mid(pstrUnicode, llngIndex, 1))
		lstrAscii = lstrUnicode & ChrB(llngAscii)
	Next
	UnicodeToAscii = lstrAscii
End Function
```


<div align="center">

## CStrCat v\. 1\.1


</div>

### Description

Faster version of StrCat posted by someone else. About 2x faster. About 6x faster than '& (ampersand)' operator for many concatenations.
 
### More Info
 
Does not need length property as it will increase the array size by 10% if it maxes out.

.add strPart

dim sc

set sc=new CStrCat

for i=0 to 10000

sc.add "whatever"

next

response.write sc.value

.value ~ whole string


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Neil McGuigan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/neil-mcguigan.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Strings](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/strings__4-26.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/neil-mcguigan-cstrcat-v-1-1__4-6746/archive/master.zip)

### API Declarations

Copyright (c) 2001, Neil McGuigan. All rights reserved. This software is licensed.


### Source Code

```
<%
Class CStrCat //v1.1
	Private i,sa()
	Public Property Get Value
		redim preserve sa(i)
		Value=Join(sa,"")
	End Property
	Private Sub Class_Initialize()
		i=clng(0)
		redim sa(500)
	End Sub
	private sub class_terminate()
		erase sa
	end sub
	Public function Add(ps)
		if len(ps)=0 then exit function
		if (i>=ubound(sa)) then upsize
		sa(i)=ps
		i=i+1
	End function
	private sub upsize()
		dim u
		u=ubound(sa)
		redim preserve sa(clng(u+u*0.1))
	end sub
End Class
%>
```


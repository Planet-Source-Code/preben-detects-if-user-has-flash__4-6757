<div align="center">

## Detects if user has Flash


</div>

### Description

This code will detect is the user has the flash plug in. And then redirect to a page with flash or one without. NICE!! please vote :)
 
### More Info
 
If users have / dont have flash

noone I think :)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Preben](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/preben.md)
**Level**          |Beginner
**User Rating**    |4.6 (41 globes from 9 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__4-1.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/preben-detects-if-user-has-flash__4-6757/archive/master.zip)

### API Declarations

Use and abuse :)


### Source Code

```
<HTML>
<HEAD>
<TITLE>www.PrebenMusic.com</TITLE>
<meta name="GENERATOR" content="PrebenMusic scripts">
</HEAD>
<BODY BGCOLOR=#000000 TEXT=FFFFFF LINK=FFFFFF>
<SCRIPT LANGUAGE="JavaScript">
<!--
var useFlash = navigator.mimeTypes &&
navigator.mimeTypes["application/x-shockwave-flash"] &&
 navigator.mimeTypes["application/x-shockwave-flash"].enabledPlugin;
//-->
</SCRIPT>
<SCRIPT LANGUAGE="VBScript">
<!--
On error resume next
useFlash = NOT IsNull(CreateObject("ShockwaveFlash.ShockwaveFlash"))
-->
</SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
<!--
if ( useFlash ) {
 window.location = "pagewithflash.htm";
} else {
 window.location = "otherpage.htm";
}
//-->
</SCRIPT>
</BODY>
</HTML>
```


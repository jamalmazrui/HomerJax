Option Explicit

Dim oJax
Dim sFile, sText, sResult

sFile = "c:\homer\HomerJax.wsc"
' Set oJax = GetObject("script:" & sFile)
Set oJax = CreateObject("Homer.Jax")
sText = "&amp;"
sResult = oJax.HtmlDecodeString(sText)
wscript.echo sResult

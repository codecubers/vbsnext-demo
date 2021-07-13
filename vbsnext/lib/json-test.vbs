' Include("classes\MyFSO.vbs")
Include("classes\VbsJson")
Dim json, str, o, i, k

Set json = New VbsJson
str="{""keys"":[1,""a""]}"
Set o = json.Decode(str)
For Each k In o("keys")
    WScript.Echo k
Next


' Set cfs = new MyFSO
' Dim f: f = putil.Resolve(".\data.json")
' EchoX ".\data.json File resolved to: %x", f
str = cfs.ReadFile(".\data\data.json")
Set o = json.Decode(str)
WScript.Echo o("Image")("Width")
WScript.Echo o("Image")("Height")
WScript.Echo o("Image")("Title")
WScript.Echo o("Image")("Thumbnail")("Url")
For Each i In o("Image")("IDs")
    WScript.Echo i
Next
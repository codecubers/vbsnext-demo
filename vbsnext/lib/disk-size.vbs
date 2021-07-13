Import "vbs-diskutil"
Dim du: set du = new DiskUtil
WScript.Echo du.GetFreeSpace("c:\")
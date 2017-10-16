*============================================================
* Removes the security warning for all files in a directory
* that match the pattern
*============================================================
LParameter tcPattern

  Local laDir[1], lnFile
  Declare Long DeleteFile in Win32Api String
  For lnFile=1 to ADir(laDir,m.tcPattern)
    DeleteFile( ;
      Addbs(FullPath(JustPath(m.tcPattern)))+;
      laDir[m.lnFile,1]+":Zone.Identifier" ;
  )
  EndFor

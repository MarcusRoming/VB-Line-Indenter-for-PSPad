VB-Line Indenter for PSPad Editor
=================================

Very simple VB-Script code beautifier for PSPad: http://www.pspad.com/ Using some code of "gogogadgetscott". 

Installation: Copy into the PSPad installation path into the folder "\Script\VBScript\". Only copy the "VBIndent.vbs" script file! 
The function can then be found via the Script-menu: "Script/Format code/VBScriptIndent". If not available activate WSH Scripting in the PSPad settings.

Known Issues:
- Cannot handle ")Then" or ""Then" in "If Then"-insctructions. A space character before the "Then" is always needed! 
- Sometimes still problems with multi-line commands.
- Somtimes problems with the colon (":") when used as a statement separator to have more than two one statements in one line. 
  This is a very rare thing since this feature is not widely known and not used very often and most uses are not problematic at all
  (for example "Dim intVar : intVar = 5" will make no trouble) . 

Use at your own risk!

New:
- Now asks if real tabs or spaces should be used.
- Corrected error counting.
- Most errors are now send to the log window instead of MsgBox.

https://github.com/MarcusRoming

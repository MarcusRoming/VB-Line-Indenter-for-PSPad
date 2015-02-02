VB-Line Indenter for PSPad Editor
=================================

Very simple VB-Script code beautifier for PSPad. Using some code of "gogogadgetscott". 

Installation: Copy into the PSPad installation path into the folder "\Script\VBScript\". Only copy the "VBIndent.vbs" script file! 
The function can then be found via the Script-menu: "Script/Format code/VBIndent". If not available activate WSH Scripting in the PSPad settings.

Fixed Bugs:
- difficulties to interpret multi line commands (using "_")
- Problems with multiple spaces in some expressions (like in "End  Function"). 

Known Issues:
- Cannot handle ")Then" or ""Then" in "If Then"-insctructions. A space character before the "Then" is always needed! 
- Sometimes still problems with multi-line commands.

Use at your own risk!

https://github.com/MarcusRoming
# AutoUnpackVBS
Autounpack for rar-files and optional cleanup

AutoUnpackVBS is a VBScript that digs down the folder structure starting with supplied path and searches for:
part01.rar, part01.rar and *.rar. It unpacks those files only and the, if the user wants, deletes them. Keeping only the unpacked file.

### Usage
AutoUnpackVBS.vbs Path DeleteAfter

### Example:
AutoUnpackVBS.vbs "C:\temp\rardirectory\" DeleteAfter

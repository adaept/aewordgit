Attribute VB_Name = "basChangeLogaewordgit"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

'=============================================================================================================================
' Tasks:
' #010 -
' #009 -
' #008 -
' #006 - Can't execute code in break mode - error after doc saved from template and opened. Use error trapping in ThisDocument
' #004 - Add an About setion in Ambigram tab to show version and logo
' #003 - Word 2019 Preview does not show 2016 Ambigram ribbon tab, report bug to Avenius
'=============================================================================================================================
'
'
' 20190608 - v003
    ' FIXED - #007 - Compile error for x64, needs PtrSafe
' 20180920 - v002
    ' FIXED - #005 - Zoom full screen and page for dotm and new doc
' 20180909 - v001 - FIXED - #001 - Implement simple test for dropdown code
    ' FIXED - #002 - Change project name to ambigram and export to .\src
' 20180903 - v000 - Use aexlgitClass as starting model for aewordgitClass



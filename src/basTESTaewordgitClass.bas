Attribute VB_Name = "basTESTaewordgitClass"
Option Explicit
Option Compare Text
Option Private Module

' Default Usage:
' The following folders are used if no custom configuration is provided:
' aexlgitType.SourceFolder = "C:\ae\adaept.com\aewordgit\src\"
' Run in immediate window:                  EXPORT_THE_CODE
' Show debug output in immediate window:    Uncomment aewordgitClassTest varDebug:="varDebug"
'
' Custom Usage:
' Public Const FOLDER_FOR_VBA_PROJECT_FILES = "Z:\The\Source\Folder\srx.MYPROJECT\"
' For custom configuration of the output source folder in aewordgitClassTest use:
' oDbObjects.SourceFolder = FOLDER_FOR_VBA_PROJECT_FILES
' Run in immediate window: EXPORT_THE_CODE
'

Public Function EXPORT_THE_CODE() As Boolean
    On Error GoTo 0
    'aewordgitClassTest
    aewordgitClassTest varDebug:="varDebug"
End Function

Public Function aewordgitClassTest(Optional ByVal varDebug As Variant, _
    Optional ByVal varSrcFldr As Variant, _
    Optional ByVal varXmlFldr As Variant, _
    Optional ByVal varXmlData As Variant) As Boolean

    On Error GoTo PROC_ERR

    Dim oXlObjects As aewordgitClass
    Set oXlObjects = New aewordgitClass

    Dim bln1 As Boolean

    If Not IsMissing(varSrcFldr) Then oXlObjects.SourceFolder = varSrcFldr      ' THE_SOURCE_FOLDER
    '''If Not IsMissing(varXmlFldr) Then oXlObjects.XMLFolder = varXmlFldr      ' THE_XML_FOLDER

Test1:
    '=============
    ' TEST 1
    '=============
    Debug.Print
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "1. aewordgitClassTest => DocumentTheWordCode"
    Debug.Print "aewordgitClassTest"
    If IsMissing(varDebug) Then
        Debug.Print , "varDebug IS missing so no parameter is passed to DocumentTheWordCode"
        Debug.Print , "DEBUGGING IS OFF"
        bln1 = oXlObjects.DocumentTheWordCode()
    Else
        Debug.Print , "varDebug IS NOT missing so blnDebug is set to True"
        bln1 = oXlObjects.DocumentTheWordCode("WithDebugging")
    End If
    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print

PROC_EXIT:
    Exit Function

PROC_ERR:
    If Err = 6068 Then ' VBA Project Not Trusted - "Programmatic access to the Visual Basic Project is not trusted..."
        MsgBox "VBA Project Not Trusted", vbCritical, "aewordgitClassTest"
        Stop
        'Resume PROC_EXIT
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aewordgitClassTest of Module basTESTaewordgitClass"
        Resume PROC_EXIT
    End If

End Function



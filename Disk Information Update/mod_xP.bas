Attribute VB_Name = "mod_xP"
Option Explicit
'=========================================================='
'Thanks to: vbAccelerator www.vbacceletaror.com            '
'Date     : 25-06-2004                                     '
'Name     : mod_xP.bas                                     '
'=========================================================='
'Daniel PC (Daniel Carrasco Olguin)                        '
'Santiago de Chile                                         '
'=========================================================='
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (Iccex As tagInitCommonControlsEx) As Boolean
Public Const ICC_USEREX_CLASSES = &H200

Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Public Sub Main()
On Error Resume Next

Dim Iccex As tagInitCommonControlsEx

With Iccex
       .lngSize = LenB(Iccex)
       .lngICC = ICC_USEREX_CLASSES
End With

InitCommonControlsEx Iccex

On Error GoTo 0
frmDisk.Show
End Sub

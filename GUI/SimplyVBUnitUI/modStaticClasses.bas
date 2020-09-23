Attribute VB_Name = "modStaticClasses"
'    CopyRight (c) 2006 Kelly Ethridge
'
'    This file is part of SimplyVBUnitUI.
'
'    SimplyVBUnitUI is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    SimplyVBUnitUI is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: modStaticClasses
'

Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Const WM_SETREDRAW As Long = &HB

Public Const IMG_NOTRUN     As String = "NotRun"
Public Const IMG_PASSED     As String = "Passed"
Public Const IMG_FAILED     As String = "Failed"
Public Const IMG_IGNORED    As String = "Ignored"


Public UserEvents As UserEvents


Public Sub SuspendLayout(ByVal hwnd As Long)
    Call SendMessage(hwnd, WM_SETREDRAW, 0, 0)
End Sub

Public Sub ResumeLayout(ByVal hwnd As Long)
    Call SendMessage(hwnd, WM_SETREDRAW, 1, 0)
End Sub

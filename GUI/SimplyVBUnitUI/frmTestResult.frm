VERSION 5.00
Begin VB.Form frmTestResult 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Test Result"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblDescription 
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Description:"
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   1320
      Width           =   945
   End
   Begin VB.Label Label6 
      Caption         =   "Message:"
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   960
      Width           =   945
   End
   Begin VB.Label lblMessage 
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label lblType 
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "Type:"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   2040
      Width           =   945
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label lblTime 
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Time:"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "Status:"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   945
   End
   Begin VB.Label lblName 
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Name:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   945
   End
End
Attribute VB_Name = "frmTestResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'    Module: frmTestResult
'

Option Explicit

Friend Sub ShowResult(ByVal Result As ITestResult)
    If Not Result Is Nothing Then
        lblName.Caption = Result.TestName
        lblStatus.Caption = GetStatus(Result)
        lblMessage.Caption = Result.Message
        lblDescription.Caption = Result.Description
        lblTime.Caption = Result.Time & " ms"
        lblType.Caption = GetType(Result)
    Else
        lblName.Caption = "Unknown"
        lblMessage.Caption = ""
        lblDescription.Caption = ""
        lblTime.Caption = ""
        lblType.Caption = ""
        lblStatus.Caption = "Not Run"
    End If
    Call Show
    Call ZOrder
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetType(ByVal Result As ITestResult) As String
    Dim ret As String
    
    If Result.IsTestResult Then
        ret = "Test"
    ElseIf Result.IsTestCaseResult Then
        ret = "Test Case"
    Else
        ret = "Test Suite"
    End If
    
    GetType = ret
End Function

Private Function GetStatus(ByVal Result As ITestResult) As String
    Dim ret As String
    
    If Not Result.Executed Then
        ret = "Not Run"
    ElseIf Result.IsIgnored Then
        ret = "Ignored"
    ElseIf Result.IsFailure Then
        ret = "Failure"
    ElseIf Result.IsTestCaseResult Then
        Dim res As TestCaseResult
        Set res = Result
        If res.HasError Or res.IsError Then
            ret = "Sub-Tests have errors"
        Else
            ret = "Success"
        End If
    Else
        ret = "Success"
    End If
    
    GetStatus = ret
End Function

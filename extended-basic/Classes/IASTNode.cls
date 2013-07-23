VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IASTNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Codegen(ByVal C As Context) As Long
    'on codegen
End Function

Public Function GetNodeType() As NODETYPE
    'on create
End Function

Public Function GetDescType(ByVal C As Context) As TypeNode
    'on codegen
End Function

Public Function IsConstant(ByVal C As Context) As Boolean
    'on codegen
End Function

Public Function Reverse(ByVal TabNum As Long) As String
    'on create
End Function

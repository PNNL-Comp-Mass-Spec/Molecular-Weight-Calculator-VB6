Attribute VB_Name = "modMwtWinDllTest"
Option Explicit

' Molecular Weight Calculator Dll test program
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)

Public objMwtWin As MwtWinDll.MolecularWeightCalculator
Public Declare Function GetTickCount Lib "kernel32" () As Long

Sub Main()
    Set objMwtWin = New MolecularWeightCalculator
    frmMwtWinDllTest.Show
    
End Sub


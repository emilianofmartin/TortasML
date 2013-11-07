VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCrear 
      Caption         =   "Crear"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TortasRicas As Collection
Private Const Ingredientes = "dddfffcccm"

Private Sub cmdCrear_Click()
    Dim Molde(2, 2) As String
    Dim e0 As String
    Dim i As Long, iLin As Integer, iCol As Integer
    
    Dim iFreeFile As Integer
    Set TortasRicas = New Collection
    
    CrearCombinacion 0, Molde, Ingredientes
    
    iFreeFile = FreeFile
    Open App.Path & "\TortasRicas.txt" For Output As #iFreeFile
    Print #iFreeFile, "Hay " & TortasRicas.Count & " tortas ricas."
    
    For i = 1 To TortasRicas.Count
        Print #iFreeFile, i & ":"
        For iLin = 0 To 2
            Print #iFreeFile, "   " & Mid$(TortasRicas(i), iLin * 3 + 1, 3)
        Next
        
        Print #iFreeFile, ""
    Next
    
    Close #iFreeFile
    
    MsgBox "Hay " & TortasRicas.Count & " tortas ricas."
End Sub

Private Sub CrearCombinacion(ByVal iStart As Integer, Mold() As String, Ingr As String)
    Dim tIngr As String, s As String, e0 As String
    Dim iLin As Integer, iCol As Integer
    
    For iLin = 0 To 2
        For iCol = 0 To 2
            If Len(Mold(iLin, iCol)) = 0 Then
                Mold(iLin, iCol) = Mid(Ingr, iStart + 1, 1)
                
                If iStart = 0 Then
                    tIngr = Mid$(Ingr, 2)
                Else
                    tIngr = Left$(Ingr, iStart) & Mid$(Ingr, iStart + 2)
                End If
                
                CrearCombinacion 0, Mold, tIngr
                Mold(iLin, iCol) = ""
                iCol = iCol - 1
                iStart = iStart + 1
                
                If Mid(Ingr, iStart + 1, 1) = "" Then
                    'Se acabaron los ingredientes
                    Exit Sub
                End If
            End If
        Next
    Next
    
    If iLin > 2 And iCol > 2 Then
        'Se terminó una torta!
        s = AnalizarTorta(Mold)
        If Len(s) Then
            For iLin = 0 To 2
                For iCol = 0 To 2
                    e0 = e0 & Mold(iLin, iCol)
                Next
            Next
            
            On Error Resume Next
            TortasRicas.Add e0, e0
            On Error GoTo 0
        End If
    End If
End Sub

Private Function AnalizarTorta(Mold() As String) As String
    Dim iLin As Integer, iCol As Integer
    Dim s As String
    AnalizarTorta = ""
    
    For iLin = 0 To 2
        s = Mold(iLin, 0)
        
        For iCol = 1 To 2
            If Mold(iLin, iCol) <> s And Mold(iLin, iCol) <> "m" Then Exit For
        Next
        
        If iCol = 3 Then
            'Todos los ingredientes son iguales en iLin o alguno de ellos era una masita
            AnalizarTorta = "i" & iLin
            Exit For
        End If
    Next
    
    If Len(AnalizarTorta) = 0 Then
        For iCol = 0 To 2
            s = Mold(0, iCol)
            
            For iLin = 1 To 2
                If Mold(iLin, iCol) <> s And Mold(iLin, iCol) <> "m" Then Exit For
            Next
            
            If iCol = 3 Then
                'Todos los ingredientes son iguales en iLin o alguno de ellos era una masita
                AnalizarTorta = "c" & iLin
                Exit For
            End If
        Next
    End If

End Function


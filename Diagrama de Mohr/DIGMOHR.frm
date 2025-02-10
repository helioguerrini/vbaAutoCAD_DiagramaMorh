VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DIGMOHR 
   Caption         =   "Diagrama de Mohr - Estado de Tensão no Ponto"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   OleObjectBlob   =   "DIGMOHR.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DIGMOHR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CALCULAR_Click()
'Por Hélio Guerrini Filho"
'Decclaração de variáveis"
Dim circulo As AcadCircle
Dim linha, C1, C2, C3, Eixos, LTextos As AcadLayer
Dim centroC1(0 To 2) As Double
Dim centroC2(0 To 2) As Double
Dim centroC3(0 To 2) As Double
Dim ponto As AcadPoint
Dim raioC1, raioC2, raioC3 As Double
Dim EIXO As AcadPolyline
Dim DIA As AcadLine
Dim X(0 To 11) As Double
Dim Y(0 To 8) As Double
Dim P(0 To 2) As Double
Dim Q(0 To 2) As Double
Dim W(0 To 2) As Double
Dim Z(0 To 2) As Double
Dim TEXTO As AcadText
Dim SIGMA, TAU As String
Dim ALTURA As Double
SIGMA = "SIGMA"
TAU = "TAU"
'Calculo do centro e raio da circunferência C3'
centroC3(0) = (Val(SIG01.Text) + Val(SIG02.Text)) / 2
centroC3(1) = 0
raioC3 = (((Val(SIG01.Text) - Val(SIG02.Text)) / 2) ^ 2 + (Val(TAU12.Text)) ^ 2) ^ 0.5
'Cria a Layer Circunferência3 de variável C3 e a torna ativa'
Set C3 = ThisDrawing.Layers.Add("Circunferência3")
ThisDrawing.ActiveLayer = C3
ThisDrawing.SetVariable "cecolor", "5"
'Substitui a circunferência por um ponto caso seu raio seja zero'
If raioC3 = 0 Then
Set ponto = ThisDrawing.ModelSpace.AddPoint(centroC3)
Else
Set circulo = ThisDrawing.ModelSpace.AddCircle(centroC3, raioC3)
End If
'Calculo do centro e raio da circunferência C2'
centroC2(0) = ((centroC3(0) + raioC3) + Val(SIG03.Text)) / 2
centroC2(1) = 0
raioC2 = (((centroC3(0) + raioC3) - Val(SIG03.Text)) ^ 2) ^ 0.5 / 2
'Cria a Layer Circunferência3 de variável C2 e a torna ativa'
Set C2 = ThisDrawing.Layers.Add("Circunferência2")
ThisDrawing.ActiveLayer = C2
ThisDrawing.SetVariable "cecolor", "4"
'Substitui a circunferência por um ponto caso seu raio seja zero'
If raioC2 = 0 Then
Set ponto = ThisDrawing.ModelSpace.AddPoint(centroC2)
Else
Set circulo = ThisDrawing.ModelSpace.AddCircle(centroC2, raioC2)
End If
'Calculo do centro e raio da circunferência C1'
centroC1(0) = ((centroC3(0) - raioC3) + Val(SIG03.Text)) / 2
centroC2(1) = 0
raioC1 = (((centroC3(0) - raioC3) - Val(SIG03.Text)) ^ 2) ^ 0.5 / 2
'Cria a Layer Circunferência3 de variável C1 e a torna ativa'
Set C1 = ThisDrawing.Layers.Add("Circunferência1")
ThisDrawing.ActiveLayer = C1
ThisDrawing.SetVariable "cecolor", "1"
'Substitui a circunferência por um ponto caso seu raio seja zero'
If raioC1 = 0 Then
Set ponto = ThisDrawing.ModelSpace.AddPoint(centroC1)
Else
Set circulo = ThisDrawing.ModelSpace.AddCircle(centroC1, raioC1)
End If
'Torna a cor branca para as layers correntes'
ThisDrawing.SetVariable "cecolor", "0"
'vetores para a criação dos eixos em polyline'
X(0) = 0
X(1) = 0
X(3) = Val(SIG03.Text)
X(4) = 0
X(6) = ((Val(SIG01.Text) + Val(SIG02.Text)) / 2) - (((Val(SIG01.Text) - Val(SIG02.Text)) / 2) ^ 2 + (Val(TAU12.Text)) ^ 2) ^ 0.5
X(7) = 0
X(9) = ((Val(SIG01.Text) + Val(SIG02.Text)) / 2) + (((Val(SIG01.Text) - Val(SIG02.Text)) / 2) ^ 2 + (Val(TAU12.Text)) ^ 2) ^ 0.5
X(10) = 0
Y(0) = 0
Y(1) = (0.05 * (raioC1 + raioC2 + raioC3)) + (raioC1 + raioC2 + raioC3) / 2
Y(3) = 0
Y(4) = 0
Y(6) = 0
Y(7) = -(0.05 * (raioC1 + raioC2 + raioC3)) - (raioC1 + raioC2 + raioC3) / 2
W(0) = 0
W(1) = (0.05 * (raioC1 + raioC2 + raioC3)) + (raioC1 + raioC2 + raioC3) / 2
If X(9) > X(6) Then
Z(0) = X(9) + 0.05 * X(9)
Else
Z(0) = X(6) + 0.05 * X(6)
End If
If X(3) > Z(0) Then
Z(0) = X(3) + 0.05 * X(3)
Else
Z(0) = Z(0)
End If
Z(1) = 0
ALTURA = 0.05 * W(1)
'cria eixos padrões para estados de tensões nulos'
If W(1) = 0 Then
X(0) = -0.2
X(1) = 0
X(3) = 1
X(4) = 0
Y(0) = 0
Y(1) = 1
Y(3) = 0
Y(4) = -0.2
W(0) = 0
W(1) = 1
Z(0) = 1
Z(1) = 0
ALTURA = 0.05
End If

P(0) = Val(SIG01.Text)
P(1) = -(Val(TAU12.Text))
Q(0) = Val(SIG02.Text)
Q(1) = Val(TAU12.Text)
'cria o layer dos eixos'
Set Eixos = ThisDrawing.Layers.Add("Eixos")
ThisDrawing.ActiveLayer = Eixos
'desenha os eixos'
Set EIXO = ThisDrawing.ModelSpace.AddPolyline(X)
Set EIXO = ThisDrawing.ModelSpace.AddPolyline(Y)
ThisDrawing.SetVariable "cecolor", "155"
Set DIA = ThisDrawing.ModelSpace.AddLine(P, Q)
'cria a layer dos textos na cor amarela'
Set LTextos = ThisDrawing.Layers.Add("Textos")
ThisDrawing.SetVariable "cecolor", "2"
ThisDrawing.ActiveLayer = LTextos
'escreve os textos'
Set TEXTO = ThisDrawing.ModelSpace.AddText(TAU, W, ALTURA)
Set TEXTO = ThisDrawing.ModelSpace.AddText(SIGMA, Z, ALTURA)
Set TEXTO = ThisDrawing.ModelSpace.AddText("(Sig1," & "-T12)", P, ALTURA)
Set TEXTO = ThisDrawing.ModelSpace.AddText("(Sig2," & "T12)", Q, ALTURA)
'Torna a cor branca para as layers correntes'
ThisDrawing.SetVariable "cecolor", "0"
'zoom em toda a extensão do desenho'
ZoomExtents
End Sub

Private Sub LIMPAR_Click()
'limpa todas as caixas de texto'
SIG01.Text = Empty
SIG02.Text = Empty
SIG03.Text = Empty
TAU12.Text = Empty
TAU21.Text = Empty
End Sub

Private Sub SIG01_Change()

End Sub

Private Sub TAU12_Change()
TAU21.Text = Val(TAU12.Text)
End Sub

Private Sub TAU21_Change()
 
End Sub

Private Sub UserForm_Click()

End Sub

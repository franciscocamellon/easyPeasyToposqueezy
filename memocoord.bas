'Macro para c�lculo de �rea e per�metro
'Autor: Sgt Fl�vio
'Data:  05 Jun 00
'Sintaxe na linha de comando: macro calculoarea
'Ultima revisao: 26 Set 00
'==============================================


'Declara��o de constantes para os quadrantes topo.
'=================================================

Const crlf$ = chr$(13) + chr$(10)
Const N = 0
Const NE = 1
Const E = 2
Const SE = 3
Const S = 4
Const SO = 5
Const O = 6
Const NO = 7
'=================================================


Type DadosdoPonto
    Nome   As String
    CoordE As Double
    CoordN As Double
End Type


'Declara��o de vari�veis a n�vel de macro

Private Ponto() As DadosdoPonto
Private GCm   As Integer
Private MCm   As Integer
Private SgCm  As Integer
Private CmSinal As Integer
Private GDm   As integer
Private MDm   As Integer
Private SgDm  As Integer
Private DmSinal As Integer
Private Kapa0 As Double
Private MCF   As String
Private Nomeimovel As String
Private NomeProp As String
Private Comarca As String
Private UF As String
Private Sist_Coord As String
Private Est_RBMC As String
Private MCSig As String
'Private GMD As Integer


Private RP() As Double
Private RPsex() As String
Private RPsexcQ() As String
Private RVerd () As Double
Private RVerdsex() As String
Private RVerdsexcQ( ) As String
Private RMagnet () As Double
Private RMagnetsex() As String
Private RMagnetsexcQ() As String
Private AzP() As Double
Private AzPsex() As String
Private AzV() As Double
Private AzVsex() As String
Private AzMag () As Double
Private AzMagsex() As String
Private QT() As Integer
Private Dist() As Double
Private Distc() As String
Private DistsKapa() As Double
Private AngH() As Double
Private AngHsex () As String
Private Cm As Double
Private Dm As Double
Private AreaTotal As Double
Private Perimetro As Double
Private Confrontantes() As String

Public UtmE As String
Public UtmN As String

Sub Main

    Dim Arquivo As String
    Dim Sugest  As String

    Filtro$ = "*.txt"
    Titulo$ = "Abrir arquivo de coordenadas"
    EsteArq$ = MbeDgnInfo.dgnFileName
    Diretorio$ = FileParse$ (EsteArq$,2)


    MbeWritePrompt "Macro Calcula Area iniciada...."


    'Abertura, leitura e armaz. dos dados do arquivo que cont�m os dados para o c�lculo (no formato): Nome, CoordE, CoordN,
    '=======================================================================================================================

    retornobotao = MbeFileOpen (Arquivo, sugest, filtro, diretorio, titulo)

        If retornobotao <> MBE_Success Then
            MbeWriteError "Macro Terminada"
            End
        End If

    Open Arquivo For Input Access Read As #1


    Contador = 0

    Do while Not Eof(1)
        Redim Preserve Ponto(Contador)
        Input #1, Ponto(Contador).Nome, Ponto(Contador).CoordE, Ponto(Contador).CoordN
        Contador = Contador + 1
    Loop


    'Cx de di�logo personalizada para ingresso da CM e DM, influentes nos c�lculos
    '==============================================================================

    'Atribui��o de valores �s vari�veis do Di�logo

    nomeimovel$ = ""
    nomeprop$ = ""
    comarca$ = ""
    GCm         = 0
    MCm         = 0
    SgCm        = 0
    GDm         = 0
    MDm         = 0
    SgDm        = 0
    Kapa0#       = 1.0
    MCF$         = ""
    CmSinal      = 1
    DmSinal      = 1
    MCSinal      = 1
    'GMD          = 1
    UfComarca% = 1
    SistCoord% = 1
    RBMC% = 1

    botao = MbeOpenModalDialog(1)

        Select Case UfComarca
            Case 1
                UF$ = "AC"
            Case 2
                UF$ = "AL"
            Case 3
                UF$ = "AP"
            Case 4
                UF$ = "AM"
            Case 5
                UF$ = "BA"
            Case 6
                UF$ = "CE"
            Case 7
                UF$ = "DF"
            Case 8
                UF$ = "ES"
            Case 9
                UF$ = "GO"
            Case 10
                UF$ = "MA"
            Case 11
                UF$ = "MT"
            Case 12
                UF$ = "MS"
            Case 13
                UF$ = "MG"
            Case 14
                UF$ = "PA"
            Case 15
                UF$ = "PB"
            Case 16
                UF$ = "PR"
            Case 17
                UF$ = "PE"
            Case 18
                UF$ = "PI"
            Case 19
                UF$ = "RJ"
            Case 20
                UF$ = "RN"
            Case 21
                UF$ = "RS"
            Case 22
                UF$ = "RO"
            Case 23
                UF$ = "RR"
            Case 24
                UF$ = "SC"
            Case 25
                UF$ = "SP"
            Case 26
                UF$ = "SE"
            Case 27
                UF$ = "TO"
        End Select

        Select Case SistCoord
            Case 1
                Sist_Coord$ = "SIRGAS 2000"
            Case 2
                Sist_Coord$ = "WGS84"
            Case 3
                Sist_Coord$ = "SAD-69"
            Case 4
                Sist_Coord$ = "C�rrego Alegre"
        End Select

        Select Case RBMC
            Case 1
                Est_RBMC$ = "RIOD"
                UtmE = "673.825,217"
                UtmN = "7.475.648,024"
            Case 2
                Est_RBMC$ = "ONRJ"
                UtmE = "682.133,192"
                UtmN = "7.466.927,822"
        End Select

        If MCSinal = 1 then
            MCsig = "EGr"
        Else
            MCsig = "WGr"
        End If

    Call Calculos

End Sub
'Function Rbmc
    'Public UtmE As Double
    'Public UtmN As Double

    'If Est_RBMC$ = "RIOD" Then
        'UtmE = 673.825,217
        'UtmN =



'End Function


Sub Calculos

    'Convers�o da Converg�ncia e Declina��o de sexagesimal p/ radianos

    Cm# = SexRad$ (GCm, MCm, SgCm, CmSinal)
    Dm# = SexRad$ (GDm, MDm, SgDm, DmSinal)


    'C�lculo do Rumos, Azimutes, Dist�ncias e Ang. horizontais
    '==============================================================================

For i = LBound(Ponto) to UBound(Ponto)

        If i<(UBound(Ponto)) Then
            PtAe# = Ponto(i).CoordE
            PtAn# = Ponto(i).CoordN
            PtBe# = Ponto(i+1).CoordE
            PtBn# = Ponto(i+1).CoordN
        Else
            PtAe# = Ponto(UBound(Ponto)).CoordE
            PtAn# = Ponto(UBound(Ponto)).CoordN
            PtBe# = Ponto(LBound(Ponto)).CoordE
            PtBn# = Ponto(LBound(Ponto)).CoordN
        End If

    'Rumo Plano em fun��o das coordenadas
    Redim Preserve RP#(i)
    RP#(i) = Rumo (PtAe#, PtAn#, PtBe#, PtBn#)

    'Rumo Plano no sist. sexagesimal
    Redim Preserve RPsex$(i)
    RPsex$(i) = RadSex$(Abs(RP#(i))) 'rumo plano sem sinal

    'Quadrante Topogr�fico em fun��o das coordenadas
    Redim Preserve QT%(i)
    QT%(i) = Quad (PtAe#, PtAn#, PtBe#, PtBn#)

    'Rumo Plano c/ sufixo indicativo do quadrante
    Redim Preserve RPsexcQ$(i)
    RPsexcQ$(i) = RPsex(i) + " " + Sufixo (QT%(i))

    'Azimute Plano em fun��o do Rumo Plano e Quadrante
    Redim Preserve AzP#(i) '(UBound(Ponto))
    AzP#(i) = Azimute (Abs(RP#(i)), QT%(i))

    'Azimute Plano no sist. sexagesimal
    Redim Preserve AzPsex$(i)
    AzPsex$(i) = RadSex$(Abs(AzP#(i)))

    'Dist�ncia em fun��o das coordenadas
    Redim Preserve DistsKapa#(i)
    DistsKapa#(i) = Distancia (PtAe#, PtAn#, PtBe#, PtBn#)
    'Dist�ncia com Kapa
    Redim Preserve Dist#(i)
    Dist#(i) = DistsKapa(i) / Kapa0
    'Dist�ncia formatada
    Redim Preserve Distc$(i)
    Distc$(i) = Format$(Str$(Dist(i)), "0.00") 'Coloca��o da dist. com 2 casas decimais
    'Dist�ncia p/ calculo de perimetro
    Mid$(Distc(i),(Instr(Distc(i),",", 0))) = "." 'substitui��o de "," por "."

    'Azimute Verdadeiro em fun��o do Azimute Plano e Cm
    Redim Preserve AzV#(i)
    AzV#(i) = AzimuteVerdad (AzP#(i), CmSinal%)

    'Azimute Verdadeiro no sist. sexagesimal
    Redim Preserve AzVsex$(i)
    AzVsex$(i) = RadSex$(AzV#(i))

    'Rumo Verdadeiro em fun��o Azimute Verdadeiro
    Redim Preserve RVerd#(i)
    RVerd#(i) = RumofAz (AzV#(i))

    'Rumo Verdadeiro no sist. sexagesimal
    Redim Preserve RVerdsex$(i)
    RVerdsex$(i) = RadSex$(RVerd#(i))

    'Rumo Verdadeiro c/ sufixo indicativo do quadrante
    Redim Preserve RVerdsexcQ$(i)
    RVerdsexcQ$(i) = RVerdsex(i) + " " + QuadranteAz (AzV#(i))

    'Azimute Magn�tico em fun��o do azimute verdadeiro e Dm
    Redim Preserve AzMag#(i)
    AzMag#(i) = AzMagnetico (AzV#(i),DmSinal%)

    'Azimute Magn�tico no sist. sexagesimal
    Redim Preserve AzMagsex$(i)
    AzMagsex$(i) = RadSex$(AzMag#(i))

    'Rumo Magn�tico em fun��o do azimute magn�tico
    Redim Preserve RMagnet#(i)
    RMagnet#(i) = RumofAz (AzMag#(i))

    'Rumo Magn�tico no sist. sexagesimal
    Redim Preserve RMagnetsex$(i)
    RMagnetsex$(i) = RadSex$(RMagnet#(i))

    'Rumo Magn�tico c/ sufixo indicativo de quadrante
    Redim Preserve RMagnetsexcQ$(i)
    RMagnetsexcQ$(i) = RMagnetsex(i) + " " + QuadranteAz (AzMag#(i))

Next i


  'C�lculo do �ngulo horizontal sistemas: decimal e sexagesimal
  For a = LBound(Ponto) to UBound(Ponto)

        If a<(UBound(Ponto)) Then
            AzPr# = AzP(a)
            AzPv# = AzP(a+1)
            Redim Preserve AngH#(UBound(Ponto))
            AngH#(a+1) = CalculaAlfa (AzPr#, AzPv#)

            Redim Preserve AngHsex$(UBound(Ponto))
            AngHsex$(a+1) = RadSex$(AngH#(a+1))
        Else
            AzPr# = AzP(UBound(Ponto))
            AzPv# = AzP(LBound(Ponto))
            AngH#(0) = CalculaAlfa (AzPr#, AzPv#)

            AngHsex$(0) = RadSex$(AngH#(0))
        End If

  Next a


    'Area
    AreaTotal# = Area#(Ponto())

    'Perimetro
    Perimetro# = Perim# (Distc$())

    Call RelatorioCalculoArea

End Sub


'================================================
'Convers�es Angulares
'================================================

'=======================
'De Radianos p/ Decimal
'=======================

Function RadDec (Rad as Double) As Double
    RadDec = 180 / PI * Rad
End Function

'==========================
'De Decimal p/ Sexagesimal
'==========================

Function DecSex (Dec#) as String
    Dim G%
    Dim M%
    Dim Sg#
    Dim PDec#

    G = Fix(Dec)
    PDec# = Dec - G
    M = Fix(PDec * 60)
    PDec# = (PDec# * 60) - M
    Sg = PDec * 60

    DecSex = Format$(Abs(G),"#0") & Chr$(176) + Format$(Abs(M),"00") & "'" & Format$(Abs(Sg),"00") & Chr$(34)

End Function

'===========================
'De Radianos p/ Sexagesimal
'===========================

Function RadSex$ (Rad#)

    RadSex = DecSex (Abs(RadDec(Rad)))

End Function


'================================================
'Convers�es de �ngulos - Revers�o
'================================================

Function SexDec#(Graus%, Min%, Seg%, Sinal%)

    Sex# = (Abs(Graus*3600) + (Min*60) + Seg)/3600
            If Sinal = 1 Then
                SexDec = Sex#
            Else
                SexDec = Sex# * (-1)
            End If
End Function

Function DecRad#(Sex#)

    DecRad = PI/180 * Sex

End Function

Function SexRad(Graus%, Min%, Seg%, Sinal%) As Double

    SexRad = DecRad(SexDec(Graus%, Min%, Seg%, Sinal%))

End Function


'================================================
'C�lculo de Rumo Plano
'================================================

Function Rumo (PtAe#, PtAn#, PtBe#, PtBn#) As Double

    Dim DE as Double
    Dim DN as Double

    DE = PtBe - PtAe
    DN = PtBn - PtAn

        If DN = 0 Then
            Rumo = PI/2
            Exit Function
        End If

    Rumo = Atn (DE / DN)

End Function

'================================================
'Determina��o de Quadrante Topogr�fico
'================================================

Function Quad (PtAe#, PtAn#, PtBe#, PtBn#) As Integer

    Dim DE as Double
    Dim DN as Double

    DE = PtBe - PtAe
    DN = PtBn - PtAn

        If DN = 0 Then
            If DE > 0 Then
                Q = E
            ElseIf DE < 0 Then
                Q = O
            End If

        ElseIf DE > 0 And DN > 0 Then
            Q = NE
        ElseIf DE > 0 And DN < 0 Then
            Q = SE
        ElseIf DE < 0 And DN < 0 Then
            Q = SO
        ElseIf DE < 0 And DN > 0 Then
            Q = NO
        ElseIf DE = 0 And DN < 0 Then
            Q = S
        ElseIf DE = 0 And DN > 0 Then
            Q = N
        End If

        Quad = Q

End Function

'================================================
'Sufixos dos rumos indicativos dos quadrantes
'================================================

Function Sufixo (QT%) As String

        If QT=0 Then
            Sufixo$ = "N"
        ElseIf QT=1 Then
            Sufixo$ = "NE"
        ElseIf QT=2 Then
            Sufixo$ = "E"
        ElseIf QT=3 Then
            Sufixo$ = "SE"
        ElseIf QT=4 Then
            Sufixo$ = "S"
        ElseIf QT=5 Then
            Sufixo$ = "SO"
        ElseIf QT=6 Then
            Sufixo$ = "O"
        Else
            Sufixo$ = "NO"
        End If
End Function

'=================================================================
'Sufixos dos rumos indicativos dos quadrantes em fun��o do azimute
'=================================================================

Function QuadranteAz (Azim#) As String

    If Azim < (PI/2) Then
        QuadranteAz = "NE"
    ElseIf Azim > (PI/2) And Azim < PI Then
        QuadranteAz = "SE"
    ElseIf Azim > PI And Azim < (3 * PI)/2 Then
        QuadranteAz = "SO"
    Else
        QuadranteAz = "NO"
    End If

End Function

'=================================================
' C�lculo de Azimute Plano
'=================================================

Function Azimute (RP#, QT%) As Double

    Dim Az As Double

    Select Case QT
        Case N
            Az = Az
        Case NE
            Az = RP
        Case E
            Az = PI/2
        Case SE
            Az = PI - RP
        Case S
            Az = PI
        Case SO
            Az = PI + RP
        Case O
            Az = (3 * PI/2)
        Case NO
            Az = (2 * PI) - RP
    End Select

            Azimute = (Az)

End Function

'================================================
'C�lculo de Dist�ncia
'================================================

Function Distancia (PtAe#, PtAn#, PtBe#, PtBn#) As Double

    Distancia = ((PtBe - PtAe)^2 + (PtBn - PtAn)^2) ^(.5)

End Function

'================================================
'C�lculo de �ngulo horizontal (topogr�fico)
'================================================

Function CalculaAlfa (AzRe#, AzVante#) As Double

        Dim Alfa As Double

        DAz# = AzVante - AzRe

        If DAz > PI Then
            Alfa = ((2*PI) - DAz) + PI
        Else
            Alfa = (PI - DAz)
        End If

        If Abs(Alfa) > (2*PI) Then
            CalculaAlfa = Alfa - (2*PI)
        Else
            CalculaAlfa = Alfa
        End If

End Function

'================================================
'C�lculo do Azimute Verdadeiro
'================================================

Function AzimuteVerdad (AzP#, CmSinal%) As Double

    If AzP > Abs(Cm) Then
        AzimuteVerdad = AzP + Cm
    Else
        'AzimuteVerdad = (2 * PI) - (Cm - AzP)
         AzimuteVerdad = Abs(Cm - AzP)
    End If

End Function

'================================================
'C�lculo do Azimute Magn�tico
'================================================

Function AzMagnetico (AzV#, DmSinal%) As Double

    If AzV > Abs(Dm) Then
        AzMagnetico = AzV - Dm
    Else
        'AzMagnetico = (2 * PI) - (Dm - AzV)
        AzMagnetico = Abs(Dm - AzV)
    End If

End Function

'================================================
'C�lculo do Rumo em fun��o do Azimute
'================================================

Function RumofAz (Azim#) As Double

    If Azim < (PI/2) Then
        RumofAz = Azim
    ElseIf Azim > (PI/2) And Azim < PI Then
        RumofAz = PI - Azim
    ElseIf Azim > PI And Azim < (3 * PI)/2 Then
        RumofAz = Azim - PI
    Else
        RumofAz = (2 * PI) - Azim
    End If

End Function

'================================================
'C�lculo de �rea
'================================================

Function Area (Ponto() As DadosdoPonto) As Double

    Dim AreaTot as Double

    AreaTot = (Ponto(0).CoordN * (Ponto(UBound(Ponto)).CoordE - Ponto(1).CoordE)) + _
              (Ponto(UBound(Ponto)).CoordN * (Ponto(UBound(Ponto)-1).CoordE - Ponto(0).CoordE))
    For i = 1 to UBound(Ponto)-1
        AreaTot = AreaTot + Ponto(i).CoordN * (Ponto(i-1).CoordE - Ponto(i+1).CoordE)
    Next i

    Area = Abs(AreaTot/2)

End Function

'================================================
'C�lculo de Per�metro
'================================================

Function Perim# (Ladosf() As String)

    Dim DisTot As Double
    Dim Lados() As Double


    For i = LBound(Ladosf) to UBound(Ladosf)
        Redim Preserve Lados(i)
        Lados(i) = Val(Ladosf(i))
        DisTot = Lados(i) + DisTot
    Next i

    Perim = DisTot

End Function

'================================================
'Rotina para impress�o do Relat�rio de C�lc. �rea
'================================================

Sub RelatorioCalculoArea

    'Vari�veis do di�logo
    Dim Filename As String
    Dim Sugest  As String


    Filtro$ = "*.are"
    Titulo$ = "Criar arquivo de C�lculo de �rea"
    EsteArq$ = MbeDgnInfo.dgnFileName
    Diretorio$ = FileParse$ (EsteArq$,2)


    'Cria��o de aquivo para abrigar o c�lculo de �rea
    '==================================================

    retornobotao = MbeFileCreate (Filename, sugest, Filtro, Diretorio, Titulo)

        If retornobotao <> MBE_Success Then
            MbeWriteError "Macro Terminada"
            End
        End If

    Open Filename For Output Access Write As #2

    'Escreve cabe�alho no arquivo

    Print #2, Space$(34) + "Minist�rio da Defesa"
    Print #2, Space$(34) + "Ex�rcito  Brasileiro"
    Print #2, Space$(25) + "Secretaria de Tecnologia da Informa��o"
    Print #2, Space$(29) + "Diretoria de Servi�o Geogr�fico"
    Print #2, Space$(33) + "3� DIV DE LEVANTAMENTO"
    Print #2,
    Print #2,
    Print #2,
    Print #2, Space$(31) + "C�LCULO DE �REA E PER�METRO"
    Print #2,
    Print #2,
    Print #2,
    Print #2, "NOME DA �REA:" + Space$(5) + Nomeimovel
    Print #2,
    Print #2,
    Print #2,

    If (GCm = 0) And (MCm = 0) And (SgCm = 0) Then
            Print #2, " " + "PONTO" + "           " + "COORD E" + "        " + "COORD N" + "   " + "  AZ PLANO" + "     " + "DIST(m)" + "       " + "RUMO PLANO" + "      " + "ANG TOPOG"
            For i = LBound(Ponto) to UBound(Ponto)
                Print #2, Ponto(i).Nome, Ponto(i).CoordE, Ponto(i).CoordN, AzPsex(i), Distc(i), RPsexcQ(i), AngHsex(i)
            Next i

    ElseIf (GDm = 0) And (MDm = 0) And (SgDm = 0) Then
            Print #2, " " + "PONTO" + "           " + "COORD E" + "        " + "COORD N" + "   " + "AZ VERDAD" + "     " + "DIST(m)" + "       " + "RUMO VERDAD" + "      " + "ANG TOPOG"
            For i = LBound(Ponto) to UBound(Ponto)
                Print #2, Ponto(i).Nome, Ponto(i).CoordE, Ponto(i).CoordN, AzVsex(i), Distc(i), RVerdsexcQ(i), AngHsex(i)
            Next i
    Else
            Print #2, " " + "PONTO" + "           " + "COORD E" + "        " + "COORD N" + "   " + "AZ VERDAD" + "     " + "DIST(m)" + "       " + "RUMO MAGNET" + "      " + "ANG TOPOG"
            For i = LBound(Ponto) to UBound(Ponto)
                Print #2, Ponto(i).Nome, Ponto(i).CoordE, Ponto(i).CoordN, AzVsex(i), Distc(i), RMagnetsexcQ(i), AngHsex(i)   ', AzMagsex(i)
            Next i

    End If

    Print #2,
    Print #2,
    Print #2,"�REA:" + Space$(10) + Format$(Str$(AreaTotal),"0.00") + "" + "m" + Chr$(178)
    Print #2,
    Print #2,"PER�METRO:" + Space$(5) + Format$(Str$(Perimetro),"0.00") + "" + "m"
    Print #2,
    Print #2,
    Print #2,"DADOS DE ENTRADA:"
    Print #2,
    If DmSinal = 2 Then
        Print #2,"DECLINA��O MAGN�TICA:" + Space$(5) + "-" + Str$(GDm) + Chr$(176) + Str$(MDm) + "'" + Str$(SgDm) + Chr$(34)
    Else
        Print #2,"DECLINA��O MAGN�TICA:" + Space$(5) + Str$(GDm) + Chr$(176) + Str$(MDm) + "'" + Str$(SgDm) + Chr$(34)
    End If

    If CmSinal = 2 Then
        Print #2,"CONVERG�NCIA MERIDIANA:" + Space$(3) + "-" + Str$(GCm) + Chr$(176) + Str$(MCm) + "'" + Str$(SgCm) + Chr$(34)
    Else
        Print #2,"CONVERG�NCIA MERIDIANA:" + Space$(3) + Str$(GCm) + Chr$(176) + Str$(MCm) + "'" + Str$(SgCm) + Chr$(34)
    End If
    Print #2,"KAPA INICIAL:"         + Space$(15) + Str$(Kapa0)

    Call ColocaValores

End Sub

'================================================
'Rotina p/cria��o de arquivo p/ coloc. de valores
'================================================

Sub ColocaValores

    'Vari�veis do di�logo
    Dim NomeArquivo As String
    Dim Sugest  As String

    Filtro$ = "*.val"
    Titulo$ = "Criar arquivo de Coloca��o de Valores"
    EsteArq$ = MbeDgnInfo.dgnFileName
    Diretorio$ = FileParse$ (EsteArq$,2)

    'Cria��o de aquivo para abrigar os valores de c�lculo p/ o desenho

    retornobotao = MbeFileCreate (NomeArquivo, sugest, Filtro, Diretorio, Titulo)

        If retornobotao <> MBE_Success Then
            MbeWriteError "Macro Terminada"
            Call GeraMemorial
        End If

    Open NomeArquivo For Output Access Write As #3

    If (GCm = 0) And (MCm = 0) And (SgCm = 0) Then
            For i = LBound(Ponto) to UBound(Ponto)
                Print #3, Ponto(i).Nome; Chr$(44); Trim$(Str$(Ponto(i).CoordE)); Chr$(44); Trim$(Str$(Ponto(i).CoordN)); Chr$(44); AzPsex(i); Chr$(44); Distc(i); Chr$(44); RPsexcQ(i); Chr$(44); AngHsex(i); Chr$(44)
            Next i

    ElseIf (GDm = 0) And (MDm = 0) And (SgDm = 0) Then
            For i = LBound(Ponto) to UBound(Ponto)
                Print #3, Ponto(i).Nome; Chr$(44); Trim$(Str$(Ponto(i).CoordE)); Chr$(44); Trim$(Str$(Ponto(i).CoordN)); Chr$(44); AzVsex(i); Chr$(44); Distc(i); Chr$(44); RVerdsexcQ(i); Chr$(44); AngHsex(i); Chr$(44)
            Next i
    Else
            For i = LBound(Ponto) to UBound(Ponto)
                Print #3, Ponto(i).Nome; Chr$(44); Trim$(Str$(Ponto(i).CoordE)); Chr$(44); Trim$(Str$(Ponto(i).CoordN)); Chr$(44); AzVsex(i); Chr$(44); Distc(i); Chr$(44); RMagnetsexcQ(i); Chr$(44); AngHsex(i); Chr$(44)
            Next i
    End If

    Call GeraMemorial

End Sub

'================================================
'Rotina p/cria��o de arquivo de memorial
'================================================

Sub GeraMemorial

    'Inicia as variav�is do di�logo Dados p/ Conf. do Mem. Descritivo

    BenfU$ = "Exemplo: xx edifica��es totalizando uma �rea coberta de xxx,xx m" + Chr$(178)
    BenfD$ = "Exemplo: xx benfeitorias totalizando uma �rea descoberta de xxx,xx m" + Chr$(178)
    Intr%  = 1
    Metod% = 1
    DescImovel$ = "Exemplo: Im�vel urbano, constitu�do por terreno e benfeitorias, localizado � (rua, av., margem de br, etc...), no munic�pio de (Recife), estado de (Pernambuco), sob a responsabilidade administrativa da(o)(3� DL)"
    Descppto$   = "Exemplo: � um marco (ou ponto, ou canto de muro, etc...) de concreto no formato tronco piramidal, nas dimens�es 0,80x0,20x0,15 m, aflorando cerca de 15 cm do solo natural, tendo em seu topo um pino de ferro (ou chapa), localizado na(rua, intersec��o, etc)"

    botao = MbeOpenModalDialog(2)

    'Atribui valores para Benfeitorias e M�todo Empregado


        Select Case Intr
            Case 1
                Instrumentos$ = "Receptor de sat�lites (Marca) (Modelo)."
            Case 2
                Instrumentos$ = "Esta��o Total (Marca) (Modelo)."
            Case 3
                Instrumentos$ = "Teodolito (Marca) (Modelo)."
            Case 4
                Instrumentos$ = "Girosc�pio (Marca) (Modelo)."
            Case 5
                Instrumentos$ = "Distanci�metro eletr�nico (Marca) (Modelo)."
            Case 6
                Instrumentos$ = "Receptor de sat�lites (Marca) (Modelo) e Esta��o Total (Marca) (Modelo)."
        End Select

        Select Case Metod
            Case 1
                Metodos$ = "Poligona��o com irradiamentos eletr�nicos."
            Case 2
                Metodos$ = "Determina��o de pontos atrav�s do Global Position Sistem (GPS), empregando o(s) processo(s): (exemplo: EST�TICO, CINEM�TICO, etc...)."
            Case 3
                Metodos$ = "Determina��o de pontos atrav�s do Global Position Sistem (GPS), empregando o(s) processo(s): (exemplo: EST�TICO, CINEM�TICO, etc...), complementado por poligonal com irradiamentos eletr�nicos."
        End Select

     'Abre cx de dial�go dados do confrontante
     Confrontante$ = ""

        For x = LBound(Ponto) to (UBound(Ponto)-1)
            PtoC$ = Ponto(x).Nome + "/" + Ponto(x+1).Nome
            botao = MbeOpenModalDialog(3)
                If botao <> 3 Then
                    End
                End If
                Redim Preserve Confrontantes(UBound(Ponto))
                Confrontantes$(x) = Confrontante
        Next x

        PtoC$ = Ponto(UBound(Ponto)).Nome + "/" + Ponto(LBound(Ponto)).Nome
        botao = MbeOpenModalDialog(3)
                If botao <> 3 Then
                    End
                End If
        Confrontantes$(UBound(Ponto)) = Confrontante


    'Vari�veis do di�logo de cria��o do arquivo memorial
    Dim NomeArq As String
    Dim Sugest  As String

    Filtro$ = "*.doc"
    Titulo$ = "Criar arquivo de Memorial Descritivo"
    EsteArq$ = MbeDgnInfo.dgnFileName
    Diretorio$ = FileParse$ (EsteArq$,2)

    'Cria��o de aquivo para abrigar os dados dol memorial

    retornobotao = MbeFileCreate (NomeArq, sugest, Filtro, Diretorio, Titulo)

        If retornobotao <> MBE_Success Then
            MbeWriteError "Macro Terminada"
            End
        End If

    Open NomeArq For Output Access Write As #4

    'Print #4, vTab + "Minist�rio da Defesa"
    'Print #4, Space$(34) + "Ex�rcito  Brasileiro"
    Print #4, Space$(32) + "F & S Topografia"
    Print #4, Space$(30) + "Divis�o de Topografia"
    Print #4, Space$(29) + "Norteando seus projetos"
    'Print #4,
    Print #4,
    Print #4,
    Print #4, Space$(31) + "MEMORIAL DESCRITIVO"
    Print #4,
    Print #4,
    Print #4, "IM�VEL: " + nomeimovel
    Print #4,
    Print #4, "Propriet�rio: " + nomeprop
    Print #4, "Comarca: " + comarca + Space$(35) + "UF: " + UF
    Print #4, "�rea (ha):" + Space$(1) + Format$(Str$(AreaTotal),"0.00") + Space$(1) + "ha" + Space$(23) + "Per�metro:" + Format$(Str$(Perimetro),"0.00") + " m"
    'Print #4,
    'Print #4, "BENFEITORIAS:" + Space$(31) + BenfU
    'Print #4, Space$(62) + BenfD
    'Print #4,
    'Print #4,
    'Print #4, "INSTRUMENTOS:" + Space$(29) + Instrumentos
    'Print #4,
    'Print #4,
    'Print #4, "M�TODO EMPREGADO:" + Space$(16) + Metodos
    Print #4,
    Print #4,
    Print #4, Space$(5) + "Inicia-se a descri��o deste per�metro no V�rtice " + Ponto(0).Nome + ", de coordenadas planas UTM E=" + Str$(Ponto(0).CoordE) + ", N=" + Str$(Ponto(0).CoordN) + ", deste, segue com azimute plano de " + AzPSex(0) + ", e dist�ncia de " + Distc(0) + " m, confrontando com "  + Confrontantes(0) + Space$(1)
    'Print #4, Descppto + ";"

    If (GCm = 0) And (MCm = 0) And (SgCm = 0) Then
            For z = (LBound(Ponto)+1) to (UBound(Ponto))
                Print #4, "encontra-se o V�rtice " + Ponto(z).Nome + ", de coordenadas planas UTM E=" + Str$(Ponto(z).CoordE) + ", N=" + Str$(Ponto(z).CoordN) + ", deste, segue com azimute plano de " + AzPSex(z) + ", e dist�ncia " + Distc(z) + " m, confrontando com "  + Confrontantes(z) + ","
            Next z
                Print #4, " deste, segue com azimute plano de " + AzPSex(UBound(Ponto)) + ", e dist�ncia " + Distc(UBound(Ponto)) + " m, confrontando com "  + Confrontantes(UBound(Ponto)) + "; encontra-se o V�rtice " + Ponto(0).Nome + ";"
    Else
            For z = (LBound(Ponto)+1) to (UBound(Ponto)-1)
                Print #4, "partindo do V�rtice " + Ponto(z).Nome + "com azimute verdadeiro de " + AzVSex(z) + ", e dist�ncia " + Distc(z) + " m, confrontando com "  + Confrontantes(z) + " encontra-se o V�rtice " + Ponto(z+1).Nome + ";"
            Next z
                Print #4, "partindo do V�rtice " + Ponto(UBound(Ponto)).Nome + "com azimute verdadeiro de " + AzVSex(UBound(Ponto)) + ", e dist�ncia " + Distc(UBound(Ponto)) + " m, confrontando com "  + Confrontantes(UBound(Ponto)) + " encontra-se o V�rtice " + Ponto(0).Nome + ";"
    End If

    Print #4, "ponto inicial da descri��o deste per�metro. Todas as coordenadas aqui descritas est�o georreferenciadas ao Sistema Geod�sico Brasileiro, a partir da esta��o ativa do IBGE (Instituto Brasileiro de Geografia e Estat�stica) de NOME: "+ Est_RBMC +" na cidade do Rio de Janeiro - RJ, de coordenadas E = "+ UtmE +" e N = "+ UtmN +", e encontram-se representadas no Sistema UTM, referenciadas ao Meridiano Central " + MCF + Space$(1) + MCSig + ", tendo como o Datum o " + Sist_Coord + ". Todos os azimutes e dist�ncias, �reas e per�metros foram calculados no plano de proje��o UTM."
    Print #4,
    Print #4,
    Print #4,
    Print #4, "Rio de Janeiro - RJ, " + Date$()
    Print #4,
    Print #4,
    Print #4,
    Print #4,
    Print #4,
    Print #4,
    Print #4, Space$(27) + "_____________________________"
    Print #4, Space$(30) + "F�bio de Souza Ananias"
    Print #4, Space$(32) + "T�cnico Agrimensor"
    Print #4, Space$(29) + "CREA n� 2004102822/TD-RJ"




End Sub

 
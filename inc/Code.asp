<%
Option Explicit
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Const nSaturation = 225  '��ɫ���
Const nBlankNoisyDotOdds = 1.0 '�հ״������(��ɫ)
Const nColorNoisyDotOdds = 1.0 '��ɫ�������(�ַ�)
Const nNoisyLineCount = 0 '����������
Const nCharMin = 4 '�������ַ���С����
Const nCharMax = 5 '�������ַ�������
Const nSpaceX = 2 '�ַ����ұ߾�
Const nSpaceY = 2 '�ַ����±߾�
Const nImgWidth = 60 '��֤���ܿ�x
Const nImgHeight = 14 '��֤���ܸ�y
Const nCharWidthRandom = 16 '�ַ������
Const nCharHeightRandom = 16 '�ַ������
Const nPositionXRandom = 10 '�ַ�������
Const nPositionYRandom = 10 '�ַ����������
Const nAngleRandom = 0 '�Ƕ������
Const nLengthRandom = 0 '���������(�ٷֱ�)
Const nColorHue = -1 'ɫ��(-2�Ҷȣ�-1��ɫ)
Const cCharSet = "0123456789"
Randomize
Dim tmp_X_1_,tmp_Y_1_,tmp_X_2_,tmp_Y_2_,tmp_n_CharC_,tmp_n_PW_,tmp_n_PH_,tmp_Pw_1_,tmp_PH_2_,str_Bur(), str_Digtal,Lines(),LineCount
tmp_n_CharC_ = nCharMin + CInt(Rnd * (nCharMax - nCharMin))
tmp_Pw_1_ = nImgWidth + 2 * nSpaceX
tmp_PH_2_ = nImgHeight + 2 * nSpaceY
call CreatValidCode("GetCode")
sub CDGen_Reset()
	LineCount = 0
	tmp_X_1_ = 0
	tmp_Y_1_ = 0
	tmp_X_2_ = 0
	tmp_Y_2_ = 1
end sub
sub CDGen_Clear()
	Dim i, j
	ReDim str_Bur(tmp_PH_2_ - 1, tmp_Pw_1_ - 1)

	for j = 0 To tmp_PH_2_ - 1
		For i = 0 To tmp_Pw_1_ - 1
			str_Bur(j, i) = 0
		next
	next
End Sub
sub Get_Position(X, Y)
	If X >= 0 and X < tmp_Pw_1_ and Y >= 0 and Y < tmp_PH_2_ Then str_Bur(Y, X) = 1
end sub
sub Get_Pline(X1, Y1, X2, Y2)
	dim DX, DY, DeltaT, i
	DX = X2 - X1
	DY = Y2 - Y1
	If Abs(DX) > Abs(DY) Then DeltaT = Abs(DX) Else DeltaT = Abs(DY)
	If DeltaT = 0 Then
		Get_Position CInt(X1),CInt(Y1)
	Else
		For i = 0 To DeltaT
			Get_Position CInt(X1 + DX * i / DeltaT), CInt(Y1 + DY * i / DeltaT)
		Next
	End If
end sub
sub CDGen_FowardDraw(nLength)
	nLength = nLength * (1 + (Rnd * 2 - 1) * nLengthRandom / 100)
	ReDim Preserve Lines(3, LineCount)
	Lines(0, LineCount) = tmp_X_1_
	Lines(1, LineCount) = tmp_Y_1_
	tmp_X_1_ = tmp_X_1_ + tmp_X_2_ * nLength
	tmp_Y_1_ = tmp_Y_1_ + tmp_Y_2_ * nLength
	Lines(2, LineCount) = tmp_X_1_
	Lines(3, LineCount) = tmp_Y_1_
	LineCount = LineCount + 1
end sub
sub CDGen_SetDirection(nAngle)
	Dim DX, DY
	nAngle = (nAngle + (Rnd * 2 - 1) * nAngleRandom) / 180 * 3.1415926
	DX = tmp_X_2_
	DY = tmp_Y_2_
	tmp_X_2_ = DX * Cos(nAngle) - DY * Sin(nAngle)
	tmp_Y_2_ = DX * Sin(nAngle) + DY * Cos(nAngle)
end sub
sub CDGen_MoveToMiddle(nActionIndex, nPercent)
	Dim DeltaX, DeltaY
	DeltaX = Lines(2, nActionIndex) - Lines(0, nActionIndex)
	DeltaY = Lines(3, nActionIndex) - Lines(1, nActionIndex)
	tmp_X_1_ = Lines(0, nActionIndex) + DeltaX * nPercent / 100
	tmp_Y_1_ = Lines(1, nActionIndex) + DeltaY * Abs(DeltaY) * nPercent / 100
end sub
sub CDGen_MoveCursor(nActionIndex)
	tmp_X_1_ = Lines(0, nActionIndex)
	tmp_Y_1_ = Lines(1, nActionIndex)
end sub
sub CDGen_Close(nActionIndex)
	ReDim Preserve Lines(3, LineCount)
	Lines(0, LineCount) = tmp_X_1_
	Lines(1, LineCount) = tmp_Y_1_
	tmp_X_1_ = Lines(0, nActionIndex)
	tmp_Y_1_ = Lines(1, nActionIndex)
	Lines(2, LineCount) = tmp_X_1_
	Lines(3, LineCount) = tmp_Y_1_
	LineCount = LineCount + 1
end sub
sub CDGen_CloseToMiddle(nActionIndex, nPercent)
	Dim DeltaX, DeltaY
	ReDim Preserve Lines(3, LineCount)
	Lines(0, LineCount) = tmp_X_1_
	Lines(1, LineCount) = tmp_Y_1_
	DeltaX = Lines(2, nActionIndex) - Lines(0, nActionIndex)
	DeltaY = Lines(3, nActionIndex) - Lines(1, nActionIndex)
	tmp_X_1_ = Lines(0, nActionIndex) + Sgn(DeltaX) * Abs(DeltaX) * nPercent / 100
	tmp_Y_1_ = Lines(1, nActionIndex) + Sgn(DeltaY) * Abs(DeltaY) * nPercent / 100
	Lines(2, LineCount) = tmp_X_1_
	Lines(3, LineCount) = tmp_Y_1_
	LineCount = LineCount + 1
end sub
sub CDGen_Flush(X0, Y0)
	Dim MaxX, MinX, MaxY, MinY
	Dim DeltaX, DeltaY, StepX, StepY, OffsetX, OffsetY
	Dim i
	MaxX = MinX = MaxY = MinY = 0
	For i = 0 To LineCount - 1
		If MaxX < Lines(0, i) Then MaxX = Lines(0, i)
		If MaxX < Lines(2, i) Then MaxX = Lines(2, i)
		If MinX > Lines(0, i) Then MinX = Lines(0, i)
		If MinX > Lines(2, i) Then MinX = Lines(2, i)
		If MaxY < Lines(1, i) Then MaxY = Lines(1, i)
		If MaxY < Lines(3, i) Then MaxY = Lines(3, i)
		If MinY > Lines(1, i) Then MinY = Lines(1, i)
		If MinY > Lines(3, i) Then MinY = Lines(3, i)
	next
	DeltaX = MaxX - MinX
	DeltaY = MaxY - MinY
	If DeltaX = 0 Then DeltaX = 1
	If DeltaY = 0 Then DeltaY = 1
	MaxX = MinX
	MaxY = MinY
	If DeltaX > DeltaY Then
		StepX = (tmp_n_PW_ - 2) / DeltaX
		StepY = (tmp_n_PH_ - 2) / DeltaX
		OffsetX = 0
		OffsetY = (DeltaX - DeltaY) / 2
	Else
		StepX = (tmp_n_PW_ - 2) / DeltaY
		StepY = (tmp_n_PH_ - 2) / DeltaY
		OffsetX = (DeltaY - DeltaX) / 2
		OffsetY = 0
	End If
	For i = 0 To LineCount - 1
		Lines(0, i) = Round((Lines(0, i) - MaxX + OffsetX) * StepX, 0)
		Lines(1, i) = Round((Lines(1, i) - MaxY + OffsetY) * StepY, 0)
		Lines(2, i) = Round((Lines(2, i) - MinX + OffsetX) * StepX, 0)
		Lines(3, i) = Round((Lines(3, i) - MinY + OffsetY) * StepY, 0)
		Get_Pline Lines(0, i) + X0 + 1, Lines(1, i) + Y0 + 1, Lines(2, i) + X0 + 1, Lines(3, i) + Y0 + 1
	Next
end sub
Sub CDGen_Char(cChar, X0, Y0)
	CDGen_Reset
	Select Case cChar
	Case "0"
		CDGen_SetDirection -60                            ' ��ʱ��60��(����ڴ�ֱ��)
		CDGen_FowardDraw -0.7                             ' ���������0.7����λ
		CDGen_SetDirection -60                            ' ��ʱ��60��
		CDGen_FowardDraw -0.7                             ' ���������0.7����λ
		CDGen_SetDirection 120                            ' ˳ʱ��120��
		CDGen_FowardDraw 1.5                              ' ����1.5����λ
		CDGen_SetDirection -60                            ' ��ʱ��60��
		CDGen_FowardDraw 0.7                              ' ����0.7����λ
		CDGen_SetDirection -60                            ' ˳ʱ��120��
		CDGen_FowardDraw 0.7                              ' ����0.7����λ
		CDGen_Close 0                                     ' ��յ�ǰ�����0��(0��ʼ)
	Case "1"
		CDGen_SetDirection -90                            ' ��ʱ��90��(����ڴ�ֱ��)
		CDGen_FowardDraw 0.5                              ' ����0.5����λ
		CDGen_MoveToMiddle 0, 50                          ' �ƶ����ʵ�λ�õ���0��(0��ʼ)��50%��
		CDGen_SetDirection 90                             ' ��ʱ��90��
		CDGen_FowardDraw -1.4                             ' ���������1.4����λ
		CDGen_SetDirection 30                             ' ��ʱ��30��
		CDGen_FowardDraw 0.4                              ' ����0.4����λ
	Case "2"
		CDGen_SetDirection 45                             ' ˳ʱ��45��(����ڴ�ֱ��)
		CDGen_FowardDraw -0.7                             ' ���������0.7����λ
		CDGen_SetDirection -120                           ' ��ʱ��120��
		CDGen_FowardDraw 0.4                              ' ����0.4����λ
		CDGen_SetDirection 60                             ' ˳ʱ��60��
		CDGen_FowardDraw 0.6                              ' ����0.6����λ
		CDGen_SetDirection 60                             ' ˳ʱ��60��
		CDGen_FowardDraw 1.6                              ' ����1.6����λ
		CDGen_SetDirection -135                           ' ��ʱ��135��
		CDGen_FowardDraw 1.0                              ' ����1.0����λ
	Case "3"
		CDGen_SetDirection -90                            ' ��ʱ��90��(����ڴ�ֱ��)
		CDGen_FowardDraw 0.8                              ' ����0.8����λ
		CDGen_SetDirection 135                            ' ˳ʱ��135��
		CDGen_FowardDraw 0.8                              ' ����0.8����λ
		CDGen_SetDirection -120                           ' ��ʱ��120��
		CDGen_FowardDraw 0.6                              ' ����0.6����λ
		CDGen_SetDirection 80                             ' ˳ʱ��80��
		CDGen_FowardDraw 0.5                              ' ����0.5����λ
		CDGen_SetDirection 60                             ' ˳ʱ��60��
		CDGen_FowardDraw 0.5                              ' ����0.5����λ
		CDGen_SetDirection 60                             ' ˳ʱ��60��
		CDGen_FowardDraw 0.5                              ' ����0.5����λ
	Case "4"
		CDGen_SetDirection 20                             ' ˳ʱ��20��(����ڴ�ֱ��)
		CDGen_FowardDraw 0.8                              ' ����0.8����λ
		CDGen_SetDirection -110                           ' ��ʱ��110��
		CDGen_FowardDraw 1.2                              ' ����1.2����λ
		CDGen_MoveToMiddle 1, 60                          ' �ƶ����ʵ�λ�õ���1��(0��ʼ)��60%��
		CDGen_SetDirection 90                             ' ˳ʱ��90��
		CDGen_FowardDraw 0.7                              ' ����0.7����λ
		CDGen_MoveCursor 2                                ' �ƶ����ʵ���2��(0��ʼ)�Ŀ�ʼ��
		CDGen_FowardDraw -1.5                             ' ���������1.5����λ
	Case "5"
		CDGen_SetDirection 90                             ' ˳ʱ��90��(����ڴ�ֱ��)
		CDGen_FowardDraw 1.0                              ' ����1.0����λ
		CDGen_SetDirection -90                            ' ��ʱ��90��
		CDGen_FowardDraw 0.8                              ' ����0.8����λ
		CDGen_SetDirection -90                            ' ��ʱ��90��
		CDGen_FowardDraw 0.8                              ' ����0.8����λ
		CDGen_SetDirection 30                             ' ˳ʱ��30��
		CDGen_FowardDraw 0.4                              ' ����0.4����λ
		CDGen_SetDirection 60                             ' ˳ʱ��60��
		CDGen_FowardDraw 0.4                              ' ����0.4����λ
		CDGen_SetDirection 30                             ' ˳ʱ��30��
		CDGen_FowardDraw 0.5                              ' ����0.5����λ
		CDGen_SetDirection 60                             ' ˳ʱ��60��
		CDGen_FowardDraw 0.8                              ' ����0.8����λ
	Case "6"
		CDGen_SetDirection -60                            ' ��ʱ��60��(����ڴ�ֱ��)
		CDGen_FowardDraw -0.7                             ' ���������0.7����λ
		CDGen_SetDirection -60                            ' ��ʱ��60��
		CDGen_FowardDraw -0.7                             ' ���������0.7����λ
		CDGen_SetDirection 120                            ' ˳ʱ��120��
		CDGen_FowardDraw 1.5                              ' ����1.5����λ
		CDGen_SetDirection 120                            ' ˳ʱ��120��
		CDGen_FowardDraw -0.7                             ' ���������0.7����λ
		CDGen_SetDirection 120                            ' ˳ʱ��120��
		CDGen_FowardDraw 0.7                              ' ����0.7����λ
		CDGen_SetDirection 120                            ' ˳ʱ��120��
		CDGen_FowardDraw -0.7                             ' ���������0.7����λ
		CDGen_SetDirection 120                            ' ˳ʱ��120��
		CDGen_FowardDraw 0.5                              ' ����0.5����λ
		CDGen_CloseToMiddle 2, 50                         ' ����ǰ����λ�����2��(0��ʼ)��50%�����
	Case "7"
		CDGen_SetDirection 180                            ' ˳ʱ��180��(����ڴ�ֱ��)
		CDGen_FowardDraw 0.3                              ' ����0.3����λ
		CDGen_SetDirection 90                             ' ˳ʱ��90��
		CDGen_FowardDraw 0.9                              ' ����0.9����λ
		CDGen_SetDirection 120                            ' ˳ʱ��120��
		CDGen_FowardDraw 1.3                              ' ����1.3����λ
	Case "8"
		CDGen_SetDirection -60                            ' ��ʱ��60��(����ڴ�ֱ��)
		CDGen_FowardDraw -0.8                             ' ���������0.8����λ
		CDGen_SetDirection -60                            ' ��ʱ��60��
		CDGen_FowardDraw -0.8                             ' ���������0.8����λ
		CDGen_SetDirection 120                            ' ˳ʱ��120��
		CDGen_FowardDraw 0.8                              ' ����0.8����λ
		CDGen_SetDirection 110                            ' ˳ʱ��110��
		CDGen_FowardDraw -1.5                             ' ���������1.5����λ
		CDGen_SetDirection -110                           ' ��ʱ��110��
		CDGen_FowardDraw 0.9                              ' ����0.9����λ
		CDGen_SetDirection 60                             ' ˳ʱ��60��
		CDGen_FowardDraw 0.8                              ' ����0.8����λ
		CDGen_SetDirection 60                             ' ˳ʱ��60��
		CDGen_FowardDraw 0.8                              ' ����0.8����λ
		CDGen_SetDirection 60                             ' ˳ʱ��60��
		CDGen_FowardDraw 0.9                              ' ����0.9����λ
		CDGen_SetDirection 70                             ' ˳ʱ��70��
		CDGen_FowardDraw 1.5	                            ' ����1.5����λ
		CDGen_Close 0                                     ' ��յ�ǰ�����0��(0��ʼ)
	Case "9"
		CDGen_SetDirection 120                            ' ��ʱ��60��(����ڴ�ֱ��)
		CDGen_FowardDraw -0.7                             ' ���������0.7����λ
		CDGen_SetDirection -60                            ' ��ʱ��60��
		CDGen_FowardDraw -0.7                             ' ���������0.7����λ
		CDGen_SetDirection -60                            ' ˳ʱ��120��
		CDGen_FowardDraw -1.5                              ' ����1.5����λ
		CDGen_SetDirection -60                            ' ˳ʱ��120��
		CDGen_FowardDraw -0.7                             ' ���������0.7����λ
		CDGen_SetDirection -60                            ' ˳ʱ��120��
		CDGen_FowardDraw -0.7                              ' ����0.7����λ
		CDGen_SetDirection 120                            ' ˳ʱ��120��
		CDGen_FowardDraw 0.7                             ' ���������0.7����λ
		CDGen_SetDirection -60                            ' ˳ʱ��120��
		CDGen_FowardDraw 0.5                              ' ����0.5����λ
		CDGen_CloseToMiddle 2, 50                         ' ����ǰ����λ�����2��(0��ʼ)��50%�����
	Case "A"
		CDGen_SetDirection -(Rnd * 20 + 150)              ' ��ʱ��150-170��(����ڴ�ֱ��)
		CDGen_FowardDraw Rnd * 0.2 + 1.1                  ' ����1.1-1.3����λ
		CDGen_SetDirection Rnd * 20 + 140                 ' ˳ʱ��140-160��
		CDGen_FowardDraw Rnd * 0.2 + 1.1                  ' ����1.1-1.3����λ
		CDGen_MoveToMiddle 0, 30                          ' �ƶ����ʵ�λ�õ���1��(0��ʼ)��30%��
		CDGen_CloseToMiddle 1, 70                         ' ����ǰ����λ�����1��(0��ʼ)��70%�����
	Case "B"
		CDGen_SetDirection -(Rnd * 20 + 50)               ' ��ʱ��50-70��(����ڴ�ֱ��)
		CDGen_FowardDraw Rnd * 0.4 + 0.8                  ' ����0.8-1.2����λ
		CDGen_SetDirection Rnd * 20 + 110                 ' ˳ʱ��110-130��
		CDGen_FowardDraw Rnd * 0.2 + 0.6                  ' ����0.6-0.8����λ
		CDGen_SetDirection -(Rnd * 20 + 110)              ' ��ʱ��110-130��
		CDGen_FowardDraw Rnd * 0.2 + 0.6                  ' ����0.6-0.8����λ
		CDGen_SetDirection Rnd * 20 + 110                 ' ˳ʱ��110-130��
		CDGen_FowardDraw Rnd * 0.4 + 0.8                  ' ����0.8-1.2����λ
		CDGen_Close 0                                     ' ��յ�ǰ�����1��(0��ʼ)
	Case "C"
		CDGen_SetDirection -60                            ' ��ʱ��60��(����ڴ�ֱ��)
		CDGen_FowardDraw -0.7                             ' ���������0.7����λ
		CDGen_SetDirection -60                            ' ��ʱ��60��

		CDGen_FowardDraw -0.7                             ' ���������0.7����λ
		CDGen_SetDirection 120                            ' ˳ʱ��120��
		CDGen_FowardDraw 1.5                              ' ����1.5����λ
		CDGen_SetDirection 120                            ' ˳ʱ��120��
		CDGen_FowardDraw -0.7                             ' ���������0.7����λ
		CDGen_SetDirection 120                            ' ˳ʱ��120��
		CDGen_FowardDraw 0.7                              ' ����0.7����λ
	Case "D"
		CDGen_SetDirection -(Rnd * 0 + 50)               ' ��ʱ��50-70��(����ڴ�ֱ��)
		CDGen_FowardDraw Rnd * 0.4 + 0.8                  ' ����0.8-1.2����λ
		CDGen_SetDirection Rnd * 0 + 110                 ' ˳ʱ��110-130��
		CDGen_FowardDraw Rnd * 0.2 + 0.6                  ' ����0.6-0.8����λ
		'CDGen_SetDirection -(Rnd * 20 + 110)              ' ��ʱ��110-130��
		'CDGen_FowardDraw Rnd * 0.2 + 0.6                  ' ����0.6-0.8����λ
		CDGen_SetDirection Rnd * 0+ 110                 ' ˳ʱ��110-130��
		CDGen_FowardDraw Rnd * 0.4 + 0.8                  ' ����0.8-1.2����λ
		CDGen_Close 0                                     ' ��յ�ǰ�����1��(0��ʼ)
	End Select
	CDGen_Flush X0, Y0
End Sub
Sub CDGen_Blur()
	Dim i, j
	For j = 1 To tmp_PH_2_ - 2
		For i = 1 To tmp_Pw_1_ - 2
			If str_Bur(j, i) = 0 Then
				If ((str_Bur(j, i - 1) Or str_Bur(j + 1, i)) and 1) <> 0 Then
					str_Bur(j, i) = 2
				End If
			End If
		Next
	Next
End Sub
Sub CDGen_NoisyLine()
	Dim i
	For i=1 To nNoisyLineCount
		Get_Pline Rnd * tmp_Pw_1_, Rnd * tmp_PH_2_, Rnd * tmp_Pw_1_, Rnd * tmp_PH_2_
	Next
End Sub
Sub Get_Dotrad()
	Dim i, j, NoisyDot, CurDot
	For j = 0 To tmp_PH_2_ - 1
		For i = 0 To tmp_Pw_1_ - 1
			If str_Bur(j, i) <> 0 Then
				If Rnd < nColorNoisyDotOdds Then
					str_Bur(j, i) = 0
				Else
					str_Bur(j, i) = nSaturation
				End If
			Else
				If Rnd < nBlankNoisyDotOdds Then
					str_Bur(j, i) = nSaturation
				Else
					str_Bur(j, i) = 0
				End If
			End If
		Next
	Next
End Sub
Sub CDGen()
	Dim i, Ch, w, x, y
	str_Digtal = ""
	CDGen_Clear
	w = nImgWidth / tmp_n_CharC_
	For i = 0 To tmp_n_CharC_ - 1
		tmp_n_PW_ = w * (1 + (Rnd * 2 - 1) * nCharWidthRandom / 100)
		tmp_n_PH_ = nImgHeight * (1 - Rnd * nCharHeightRandom / 100)
		x = nSpaceX + w * (i + (Rnd * 2 - 1) * nPositionXRandom / 100)
		y = nSpaceY + nImgHeight * (Rnd * 2 - 1) * nPositionYRandom / 100
		Ch = Mid(cCharSet, Int(Rnd * Len(cCharSet)) + 1, 1)
		str_Digtal = str_Digtal + Ch
		CDGen_Char Ch, x, y
	Next
	CDGen_Blur
	CDGen_NoisyLine
	Get_Dotrad
End Sub
Function str_HSBToRGB(vH, vS, vB)
	Dim aRGB(3), RGB1st, RGB2nd, RGB3rd
	Dim nH, nS, nB
	Dim lH, nF, nP, nQ, nT
	vH = (vH Mod 360)
	If vS > 100 Then
		vS = 100
	ElseIf vS < 0 Then
		vS = 0
	End If
	If vB > 100 Then
		vB = 100
	ElseIf vB < 0 Then
		vB = 0
	End If
	If vS > 0 Then
		nH = vH / 60
		nS = vS / 100
		nB = vB / 100
		lH = Int(nH)
		nF = nH - lH
		nP = nB * (1 - nS)
		nQ = nB * (1 - nS * nF)
		nT = nB * (1 - nS * (1 - nF))
		Select Case lH
		Case 0
			aRGB(0) = nB * 255
			aRGB(1) = nT * 255
			aRGB(2) = nP * 255
		Case 1
			aRGB(0) = nQ * 255
			aRGB(1) = nB * 255
			aRGB(2) = nP * 255
		Case 2
			aRGB(0) = nP * 255
			aRGB(1) = nB * 255
			aRGB(2) = nT * 255
		Case 3
			aRGB(0) = nP * 255
			aRGB(1) = nQ * 255
			aRGB(2) = nB * 255
		Case 4
			aRGB(0) = nT * 255
			aRGB(1) = nP * 255
			aRGB(2) = nB * 255
		Case 5
			aRGB(0) = nB * 255
			aRGB(1) = nP * 255
			aRGB(2) = nQ * 255
		End Select
	Else
		aRGB(0) = (vB * 255) / 100
		aRGB(1) = aRGB(0)
		aRGB(2) = aRGB(0)
	End If
	str_HSBToRGB = ChrB(Round(aRGB(2), 0)) & ChrB(Round(aRGB(1), 0)) & ChrB(Round(aRGB(0), 0))
End Function
Sub CreatValidCode(Foosun)
	Dim i, j, CurColorHue
	CDGen
	Session(Foosun) = lcase(str_Digtal)	'��¼��Session
	Dim FileSize, PicDataSize
	PicDataSize = tmp_Pw_1_ * tmp_PH_2_ * 3
	FileSize = PicDataSize + 54
	Response.BinaryWrite ChrB(66) & ChrB(77) & ChrB(FileSize Mod 256) & ChrB((FileSize \ 256) Mod 256) & ChrB((FileSize \ 256 \ 256) Mod 256) & ChrB(FileSize \ 256 \ 256 \ 256) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(54) & ChrB(0) & ChrB(0) & ChrB(0)
	Response.BinaryWrite ChrB(40) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(tmp_Pw_1_ Mod 256) & ChrB((tmp_Pw_1_ \ 256) Mod 256) & ChrB((tmp_Pw_1_ \ 256 \ 256) Mod 256) & ChrB(tmp_Pw_1_ \ 256 \ 256 \ 256) & ChrB(tmp_PH_2_ Mod 256) & ChrB((tmp_PH_2_ \ 256) Mod 256) & ChrB((tmp_PH_2_ \ 256 \ 256) Mod 256) & ChrB(tmp_PH_2_ \ 256 \ 256 \ 256) & ChrB(1) & ChrB(0) & ChrB(24) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(PicDataSize Mod 256) & ChrB((PicDataSize \ 256) Mod 256) & ChrB((PicDataSize \ 256 \ 256) Mod 256) & ChrB(PicDataSize \ 256 \ 256 \ 256) & ChrB(18) & ChrB(11) & ChrB(0) & ChrB(0) & ChrB(18) & ChrB(11) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0)
	If nColorHue = -1 Then
		CurColorHue = Int(Rnd * 360)
		CurColorHue = 200
	ElseIf nColorHue <> -2 Then
		CurColorHue = nColorHue
	End If
	For j = 0 To tmp_PH_2_ - 1
		For i = 0 To tmp_Pw_1_ - 1
			If nColorHue = -2 Then
				Response.BinaryWrite str_HSBToRGB(0, 0, 100 - str_Bur(tmp_PH_2_ - 1 - j, i))
			Else
				Response.BinaryWrite str_HSBToRGB(CurColorHue, str_Bur(tmp_PH_2_ - 1 - j, i), 100)
			End If
		Next
	Next
End Sub
%>
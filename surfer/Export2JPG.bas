Sub Main
'#���ܣ�������srf��ʽ�ļ���������ΪͼƬ
Dim SurferApp As Object
Set SurferApp = CreateObject("Surfer.Application")
SurferApp.Visible = True


'#��Ҫ�޸ĵĲ���

SRFPATH="D:\LLXM\���ݹ㳡\DAPC\DAP\yt_fz\"   '����srf�ļ������ļ���
SAVEPATH="D:\LLXM\���ݹ㳡\������ͼ2\"    '���õ���ͼƬ�����ļ���

s=0      '��ʼ�Ƕ�
e=350   '��ֹ�Ƕ�
sp=10   '���

CWidth=733    '����ͼƬ���
CHeight=752   '����ͼƬ�߶�
'#

'#�����޸�#
For i = s To e Step sp

	If i < 10 Then
			sn0= "00"
		ElseIf i < 100 Then
			sn0= "0"
		Else
			sn0 = ""
	End If
	sn1 = sn0 + CStr(i)

SrfFile=SRFPATH + sn1 + ".srf"
SaveFile=SAVEPATH +  sn1 + ".jpg"
OptionsString="Width=" + CStr(CWidth) + "," + "Height=" + CStr(CHeight)

Dim Plot As Object
Set Plot = SurferApp.Documents.Open(SrfFile)
Plot.Export(FileName:= SaveFile, SelectionOnly:=False , Options:=OptionsString)

Next i
'#
End Sub

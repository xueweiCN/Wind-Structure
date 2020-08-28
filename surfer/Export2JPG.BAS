Sub Main
'#功能：将已有srf格式文件批量导出为图片
Dim SurferApp As Object
Set SurferApp = CreateObject("Surfer.Application")
SurferApp.Visible = True


'#需要修改的部分

SRFPATH="D:\LLXM\海螺广场\DAPC\DAP\yt_fz\"   '设置srf文件所在文件夹
SAVEPATH="D:\LLXM\海螺广场\风振云图2\"    '设置导出图片所在文件夹

s=0      '起始角度
e=350   '终止角度
sp=10   '间隔

CWidth=733    '设置图片宽度
CHeight=752   '设置图片高度
'#

'#无需修改#
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

<%
	'------------------------------------------------------------------------------------------------------------
	' ����
	'------------------------------------------------------------------------------------------------------------
	' ���˽ð�
'	If CDate("2018-04-12 00:00:00") <= Now() And Now() < CDate("2018-04-12 07:00:00") Then
'		Response.Redirect "/index_180412.asp"
'		Response.End
'	End If

	'������ �б� ó�� S

	'�п��� ���� �������� �Լ� S 
	Function GetCampusVarsSiteMGC(VarsName)
		
		'���� ������ 
		gMGC_RequestServerName = LCase(Request.ServerVariables("SERVER_NAME"))  '������ �̸�
		gMGC_ServerName = "http://" & RequestServerName	

		gMGC_RequestPathInfo = LCase(Request.ServerVariables("PATH_INFO"))	' ���������
		
		'����Ʈ ���� ���� S
		
		'����	�����	http://gangnam.megastudy.net	CD0206	CD0001	http://mgangnam.megastudy.net
		'���� ��.����	http://gangnams.megastudy.net	CD0205	CD0001	http://mgangnams.megastudy.net
		'����	�����	http://gangdong.megastudy.net	CD0240	CD0198	http://mgangdong.megastudy.net
		'���� ��.��.��	http://gangdongs.megastudy.net	CD0215	CD0198	http://mgangdongs.megastudy.net
		'����	�����	http://gangbuk.megastudy.net	CD0210	CD0005	http://mgangbuk.megastudy.net
		'���� ����	http://gangbuks.megastudy.net	CD0209	CD0005	http://mgangbuks.megastudy.net
		'�뷮��	�����	http://noryangjin.megastudy.net	CD0211	CD0006	http://mnoryangjin.megastudy.net
		'�뷮�� �ܰ�	http://nrj.megastudy.net	CD0212	CD0006	http://mnrj.megastudy.net
		'����	�����	http://seocho.megastudy.net	CD0208	CD0004	http://mseocho.megastudy.net
		'���� ����	http://seochos.megastudy.net	CD0207	CD0004	http://mseochos.megastudy.net
		'����	�����	http://seongbuk.megastudy.net	CD0239	CD0098	http://mseongbuk.megastudy.net
		'���� ��.��.��	http://seongbuks.megastudy.net	CD0214	CD0098	http://mseongbuks.megastudy.net
		'���� ������Թ�	http://sinchon.megastudy.net	CD0213	CD0036	http://msinchon.megastudy.net
		'����	�����	http://pyeongchon.megastudy.net	CD0217	CD0179	http://mpyeongchon.megastudy.net
		'���̰���	http://pyeongchons.megastudy.net	CD0216	CD0179	http://mpyeongchons.megastudy.net
		'���� ����	������Թ�	http://yangji.megastudy.net	CD0220	CD0178	http://myangji.megastudy.net
		'���� �Ű�	������Թ�	http://yangjim.megastudy.net	CD0243	CD0039	http://myangjim.megastudy.net
        '��õ������Թ�	http://bucheon.megastudy.net	CD0243	CD0039	http://mbucheon.megastudy.net
		'���� �ް����͵� ����п� 	������Թ�	http://seochob.megastudy.net	CD0243	CD0039	http://mseochob.megastudy.net
		'���� ��.��.��	http://yangju.megastudy.net	CD0259	CD0258	http://myangju.megastudy.net
		'���� ��.��.��	http://unjung.megastudy.net	CD0261	CD0260	http://munjung.megastudy.net
		If IsDevSvr() = True Then
			cDev = "DEV-"
		Else
			cDev = ""
		End If

		If gMGC_RequestServerName = "localhost" Then	' ���ÿ��� �۾��� ���� �߰�(2020.03.10)

			If InStr(gMGC_RequestPathInfo, "/gangnam/nsu") > 0 Then

				'���� �����
				gMGC_Campus_N = "gangnam"
				gMGC_Campus_Code = "CD0001"
				gMGC_Campus_KN = "���� �ް����͵��п�"
				gMGC_Campus_EN = "gangnam"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "�����"
				gMGC_Campus_Detail_Code = "CD0206"
				gMGC_Campus_DirSite ="/gangnam/nsu"
				gMGC_Campus_Title = cDev&"���� �ް����͵��п�"
				gMGC_Campus_Theme = "red"

			ElseIf InStr(gMGC_RequestPathInfo, "/gangnam/jaehak") > 0 Then

				'���� ��.����
				gMGC_Campus_N = "gangnam"
				gMGC_Campus_Code = "CD0001"
				gMGC_Campus_KN = "���� �ް����͵��п�"
				gMGC_Campus_EN = "gangnam"
				gMGC_Campus_Kind = "jaehak"
				gMGC_Campus_Kind_KN = "����"
				gMGC_Campus_Detail_Code = "CD0205"
				gMGC_Campus_DirSite ="/gangnam/jaehak"
				gMGC_Campus_Title = cDev&"���� �ް����͵��п� �߰���"
				gMGC_Campus_Theme = "blue"

			ElseIf InStr(gMGC_RequestPathInfo, "/gangdong/nsu") > 0 Then
				
				'���� �����
				gMGC_Campus_N = "gangdong"
				gMGC_Campus_Code = "CD0198"
				gMGC_Campus_KN = "���� �ް����͵��п�"
				gMGC_Campus_EN = "gangdong"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "�����"
				gMGC_Campus_Detail_Code = "CD0240"
				gMGC_Campus_DirSite ="/gangdong/nsu"
				gMGC_Campus_Title = cDev&"���� �ް����͵��п�"
				gMGC_Campus_Theme = "red"

			ElseIf InStr(gMGC_RequestPathInfo, "/gangdong/jaehak") > 0 Then

				'���� ��.��.��
				gMGC_Campus_N = "gangdong"
				gMGC_Campus_Code = "CD0198"
				gMGC_Campus_KN = "���� �ް����͵��п�"
				gMGC_Campus_EN = "gangdong"
				gMGC_Campus_Kind = "jaehak"
				gMGC_Campus_Kind_KN = "���߰��"
				gMGC_Campus_Detail_Code = "CD0215"
				gMGC_Campus_DirSite ="/gangdong/jaehak"
				gMGC_Campus_Title = cDev&"���� �ް����͵��п� ���߰��"
				gMGC_Campus_Theme = "blue" 

			ElseIf InStr(gMGC_RequestPathInfo, "/gangbuk/nsu") > 0 Then

				'���� �����
				gMGC_Campus_N = "gangbuk"
				gMGC_Campus_Code = "CD0005"
				gMGC_Campus_KN = "���� �ް����͵��п�"
				gMGC_Campus_EN = "gangbuk"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "�����"
				gMGC_Campus_Detail_Code = "CD0210"
				gMGC_Campus_DirSite ="/gangbuk/nsu"
				gMGC_Campus_Title = cDev&"���� �ް����͵��п�"
				gMGC_Campus_Theme = "red"

			ElseIf InStr(gMGC_RequestPathInfo, "/gangbuk/jaehak") > 0 Then

				'���� ����
				gMGC_Campus_N = "gangbuk"
				gMGC_Campus_Code = "CD0005"
				gMGC_Campus_KN = "���� �ް����͵��п�"
				gMGC_Campus_EN = "gangbuk"
				gMGC_Campus_Kind = "jaehak"
				gMGC_Campus_Kind_KN = "��� ���ա��ܰ��п�"
				gMGC_Campus_Detail_Code = "CD0209"
				gMGC_Campus_DirSite ="/gangbuk/jaehak"
				gMGC_Campus_Title = cDev&"���� �ް����͵��п� ����"
				gMGC_Campus_Theme = "blue" 

			ElseIf InStr(gMGC_RequestPathInfo, "/noryangjin/nsu") > 0 Then

				'�뷮�� �����
				gMGC_Campus_N = "noryangjin"
				gMGC_Campus_Code = "CD0006"
				gMGC_Campus_KN = "�뷮�� �ް����͵��п�"
				gMGC_Campus_EN = "noryangjin"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "�����"
				gMGC_Campus_Detail_Code = "CD0211"
				gMGC_Campus_DirSite ="/noryangjin/nsu"
				gMGC_Campus_Title = cDev&"�뷮�� �ް����͵��п�"
				gMGC_Campus_Theme = "red"

			ElseIf InStr(gMGC_RequestPathInfo, "/noryangjin/danka") > 0 Then

				'�뷮�� �ܰ�
				gMGC_Campus_N = "noryangjin"
				gMGC_Campus_Code = "CD0006"
				gMGC_Campus_KN = "�뷮�� �ް����͵��п�"
				gMGC_Campus_EN = "noryangjin"
				gMGC_Campus_Kind = "danka"
				gMGC_Campus_Kind_KN = "�ܰ�"
				gMGC_Campus_Detail_Code = "CD0212"
				gMGC_Campus_DirSite ="/noryangjin/danka"
				gMGC_Campus_Title = cDev&"�뷮�� �ް����͵��п� �ܰ�"
				gMGC_Campus_Theme = "blue" 

			ElseIf InStr(gMGC_RequestPathInfo, "/seocho/nsu") > 0 Then

				'���� �����
				gMGC_Campus_N = "seocho"
				gMGC_Campus_Code = "CD0004"
				gMGC_Campus_KN = "���� �ް����͵��п�"
				gMGC_Campus_EN = "seocho"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "�����"
				gMGC_Campus_Detail_Code = "CD0208"
				gMGC_Campus_DirSite ="/seocho/nsu"
				gMGC_Campus_Title = cDev&"���� �ް����͵��п�"
				gMGC_Campus_Theme = "red"

			ElseIf InStr(gMGC_RequestPathInfo, "/seocho/jaehak") > 0 Then

				'���� ����
				gMGC_Campus_N = "seocho"
				gMGC_Campus_Code = "CD0004"
				gMGC_Campus_KN = "���� �ް����͵��п�"
				gMGC_Campus_EN = "seocho"
				gMGC_Campus_Kind = "jaehak"
				gMGC_Campus_Kind_KN = "����"
				gMGC_Campus_Detail_Code = "CD0207"
				gMGC_Campus_DirSite ="/seocho/jaehak"
				gMGC_Campus_Title = cDev&"���� �ް����͵��п� ����"
				gMGC_Campus_Theme = "blue"

			ElseIf InStr(gMGC_RequestPathInfo, "/seongbuk/nsu") > 0 Then
				
				'���� �����
				gMGC_Campus_N = "seongbuk"
				gMGC_Campus_Code = "CD0098"
				gMGC_Campus_KN = "���� �ް����͵��п�"
				gMGC_Campus_EN = "seongbuk"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "�����"
				gMGC_Campus_Detail_Code = "CD0239"
				gMGC_Campus_DirSite ="/seongbuk/nsu"
				gMGC_Campus_Title = cDev&"���� �ް����͵��п�"
				gMGC_Campus_Theme = "red" 

			ElseIf InStr(gMGC_RequestPathInfo, "/seongbuk/jaehak") > 0 Then

				'���� ��.��.��
				gMGC_Campus_N = "seongbuk"
				gMGC_Campus_Code = "CD0098"
				gMGC_Campus_KN = "���� �ް����͵��п�"
				gMGC_Campus_EN = "seongbuk"
				gMGC_Campus_Kind = "jaehak"
				gMGC_Campus_Kind_KN = "���߰��"
				gMGC_Campus_Detail_Code = "CD0214"
				gMGC_Campus_DirSite ="/seongbuk/jaehak"
				gMGC_Campus_Title = cDev&"���� �ް����͵��п� ���߰��"
				gMGC_Campus_Theme = "blue" 

			ElseIf InStr(gMGC_RequestPathInfo, "/sinchon") > 0 Then

				'���� ������Թ�
				gMGC_Campus_N = "sinchon"
				gMGC_Campus_Code = "CD0036"
				gMGC_Campus_KN = "���� �ް����͵��п�"
				gMGC_Campus_EN = "sinchon"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "���/�ܰ�"
				gMGC_Campus_Detail_Code = "CD0213"
				gMGC_Campus_DirSite ="/sinchon"
				gMGC_Campus_Title = cDev&"���� �ް����͵��п�"
				gMGC_Campus_Theme = "red"

			ElseIf InStr(gMGC_RequestPathInfo, "/pyeongchon/nsu") > 0 Then
				
				'���� �����
				gMGC_Campus_N = "pyeongchon"
				gMGC_Campus_Code = "CD0179"
				gMGC_Campus_KN = "���� �ް����͵��п�"
				gMGC_Campus_EN = "pyeongchon"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "�����"
				gMGC_Campus_Detail_Code = "CD0217"
				gMGC_Campus_DirSite ="/pyeongchon/nsu"
				gMGC_Campus_Title = cDev&"���� �ް����͵��п�"
				gMGC_Campus_Theme = "red" 

			ElseIf InStr(gMGC_RequestPathInfo, "/pyeongchon/jaehak") > 0 Then

				'���� ����
				gMGC_Campus_N = "pyeongchon"
				gMGC_Campus_Code = "CD0179"
				gMGC_Campus_KN = "���� �ް����͵��п�"
				gMGC_Campus_EN = "pyeongchon"
				gMGC_Campus_Kind = "jaehak"
				gMGC_Campus_Kind_KN = "����"
				gMGC_Campus_Detail_Code = "CD0216"
				gMGC_Campus_DirSite ="/pyeongchon/jaehak"
				gMGC_Campus_Title = cDev&"���� �ް����͵��п� ����"
				gMGC_Campus_Theme = "blue" 

			ElseIf InStr(gMGC_RequestPathInfo, "/yangji") > 0 Then

				'���� ���� ������Թ�
				gMGC_Campus_N = "yangji"
				gMGC_Campus_Code = "CD0178"
				gMGC_Campus_KN = "�ֻ����� ������ ���� ����п�"
				gMGC_Campus_EN = "yangji"
				gMGC_Campus_Kind = "gisuk"
				gMGC_Campus_Kind_KN = ""
				gMGC_Campus_Detail_Code = "CD0220"
				gMGC_Campus_DirSite ="/yangji"
				gMGC_Campus_Title = cDev&"�ֻ����� ������ ���� ����п�"
				gMGC_Campus_Theme = "beige"

			ElseIf InStr(gMGC_RequestPathInfo, "/yangjim") > 0 Then
				
				'���ʱ�� ������Թ�
				gMGC_Campus_N = "yangjim"
				gMGC_Campus_Code = "CD0039"
				gMGC_Campus_KN = "�ڿ��� ������ ���� ����п�"
				gMGC_Campus_EN = "yangjim"
				gMGC_Campus_Kind = "gisuk"
				gMGC_Campus_Kind_KN = ""
				gMGC_Campus_Detail_Code = "CD0243"
				gMGC_Campus_DirSite ="/yangjim"
				gMGC_Campus_Title = cDev&"�ڿ��� ������ ���� ����п�"
				gMGC_Campus_Theme = "beige"

			ElseIf InStr(gMGC_RequestPathInfo, "/bucheon/nsu") > 0 Then

				'��õ ������Թ�
				gMGC_Campus_N = "bucheon"
				gMGC_Campus_Code = "CD0250"
				gMGC_Campus_KN = "��õ �ް����͵��п�"
				gMGC_Campus_EN = "bucheon"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "�����"
				gMGC_Campus_Detail_Code = "CD0251"
				gMGC_Campus_DirSite ="/bucheon/nsu"
				gMGC_Campus_Title = cDev&"��õ �ް����͵��п�"
				gMGC_Campus_Theme = "red"

			ElseIf InStr(gMGC_RequestPathInfo, "/bundang/nsu") > 0 Then

				'�д� ������Թ�
				gMGC_Campus_N = "bundang"
				gMGC_Campus_Code = "CD0252"
				gMGC_Campus_KN = "�д� �ް����͵��п�"
				gMGC_Campus_EN = "bundang"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "�����"
				gMGC_Campus_Detail_Code = "CD0253"
				gMGC_Campus_DirSite ="/bundang/nsu"
				gMGC_Campus_Title = cDev&"�д� �ް����͵��п�"
				gMGC_Campus_Theme = "red" 

			ElseIf InStr(gMGC_RequestPathInfo, "/ilsan/nsu") > 0 Then

				'�ϻ� ������Թ�
				gMGC_Campus_N = "ilsan"
				gMGC_Campus_Code = "CD0254"
				gMGC_Campus_KN = "�ϻ� �ް����͵��п�"
				gMGC_Campus_EN = "ilsan"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "�����"
				gMGC_Campus_Detail_Code = "CD0255"
				gMGC_Campus_DirSite ="/ilsan/nsu"
				gMGC_Campus_Title = cDev&"�ϻ� �ް����͵��п�"
				gMGC_Campus_Theme = "red" 

			ElseIf InStr(gMGC_RequestPathInfo, "/yangju/jaehak") > 0 Then

				'���� ��.��.��
				gMGC_Campus_N = "yangju"
				gMGC_Campus_Code = "CD0258"
				gMGC_Campus_KN = "���� �ް����͵��п�"
				gMGC_Campus_EN = "yangju"
				gMGC_Campus_Kind = "jaehak"
				gMGC_Campus_Kind_KN = "���߰��"
				gMGC_Campus_Detail_Code = "CD0259"
				gMGC_Campus_DirSite ="/yangju/jaehak"
				gMGC_Campus_Title = cDev&"���� �ް����͵��п� ���߰��"
				gMGC_Campus_Theme = "blue"

			ElseIf InStr(gMGC_RequestPathInfo, "/unjung/jaehak") > 0 Then

				'���� ��.��.��
				gMGC_Campus_N = "unjung"
				gMGC_Campus_Code = "CD0260"
				gMGC_Campus_KN = "���� �ް����͵��п�"
				gMGC_Campus_EN = "unjung"
				gMGC_Campus_Kind = "jaehak"
				gMGC_Campus_Kind_KN = "�߰���"
				gMGC_Campus_Detail_Code = "CD0261"
				gMGC_Campus_DirSite ="/unjung/jaehak"
				gMGC_Campus_Title = cDev&"���� �ް����͵��п� ���߰��"
				gMGC_Campus_Theme = "blue"

			ElseIf InStr(gMGC_RequestPathInfo, "/songpa/nsu") > 0 Then

				'���� �����
				gMGC_Campus_N = "songpa"
				gMGC_Campus_Code = "CD0275"
				gMGC_Campus_KN = "���� �ް����͵��п�"
				gMGC_Campus_EN = "songpa"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "�����"
				gMGC_Campus_Detail_Code = "CD0276"
				gMGC_Campus_DirSite ="/songpa/nsu"
				gMGC_Campus_Title = cDev&"���� �ް����͵��п�"
				gMGC_Campus_Theme = "red"

			ElseIf InStr(gMGC_RequestPathInfo, "/ansung/") > 0 Then
				
				'�ȼ���� ������Թ�
				gMGC_Campus_N = "ansung"
				gMGC_Campus_Code = "CD0279"
				gMGC_Campus_KN = "�ȼ� �ް����͵� ����п�"
				gMGC_Campus_EN = "ansung"
				gMGC_Campus_Kind = "gisuk"
				gMGC_Campus_Kind_KN = ""
				gMGC_Campus_Detail_Code = "CD0280"
				gMGC_Campus_DirSite ="/ansung"
				gMGC_Campus_Title = cDev&"�ȼ� �ް����͵� ����п�"
				gMGC_Campus_Theme = "beige"

			Else
					' �� �� 
					gMGC_Campus_N = "campus"
					gMGC_Campus_Code = ""
					gMGC_Campus_KN = "�ް����͵��п�"
					gMGC_Campus_EN = "megastudy campus"
					gMGC_Campus_Kind = ""
					gMGC_Campus_Kind_KN = ""
					gMGC_Campus_Detail_Code = ""
					gMGC_Campus_DirSite ="/campus"
					gMGC_Campus_Title = cDev&"�ް����͵��п�"
					gMGC_Campus_Theme = ""

			End If

		Else	' ���� �۾� �ƴϸ� ���� ���� Ž
			
			Select Case gMGC_RequestServerName

				'Campus 
				Case "campus.megastudy.net", "campus1.megastudy.net", "campus2.megastudy.net", "campus5.megastudy.net", "amsdev.megastudy.net" ,"tcampus.megastudy.net", "mcampus.megastudy.net", "mcampus1.megastudy.net", "mcampus2.megastudy.net", "mcampus5.megastudy.net" , "tmcampus.megastudy.net" , "campus3.megastudy.net" , "mcampus3.megastudy.net"
		   
					gMGC_Campus_N = ""
					gMGC_Campus_Code = ""
					gMGC_Campus_KN = "�ް����͵��п�"
					gMGC_Campus_EN = "megastudy campus"
					gMGC_Campus_Kind = ""
					gMGC_Campus_Kind_KN = ""
					gMGC_Campus_Detail_Code = ""
					gMGC_Campus_DirSite ="/campus"
					gMGC_Campus_Title = cDev&"�ް����͵��п�"
					gMGC_Campus_Theme = ""
				
				'���� �����	http://gangnam.megastudy.net	http://mgangnam.megastudy.net
				Case "gangnam.megastudy.net", "mgangnam.megastudy.net", "tgangnam.megastudy.net", "tmgangnam.megastudy.net"

					gMGC_Campus_N = "gangnam"
					gMGC_Campus_Code = "CD0001"
					gMGC_Campus_KN = "���� �ް����͵��п�"
					gMGC_Campus_EN = "gangnam"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "�����"
					gMGC_Campus_Detail_Code = "CD0206"
					gMGC_Campus_DirSite ="/gangnam/nsu"
					gMGC_Campus_Title = cDev&"���� �ް����͵��п�"
					gMGC_Campus_Theme = "red"
				
				'���� ��.����	http://gangnams.megastudy.net	http://mgangnams.megastudy.net
				Case "gangnams.megastudy.net", "mgangnams.megastudy.net", "tgangnams.megastudy.net", "tmgangnams.megastudy.net"
					
					gMGC_Campus_N = "gangnam"
					gMGC_Campus_Code = "CD0001"
					gMGC_Campus_KN = "���� �ް����͵��п�"
					gMGC_Campus_EN = "gangnam"
					gMGC_Campus_Kind = "jaehak"
					gMGC_Campus_Kind_KN = "����"
					gMGC_Campus_Detail_Code = "CD0205"
					gMGC_Campus_DirSite ="/gangnam/jaehak"
					gMGC_Campus_Title = cDev&"���� �ް����͵��п� �߰���"
					gMGC_Campus_Theme = "blue"
					
				'���� �����	http://gangdong.megastudy.net	http://mgangdong.megastudy.net
				Case "gangdong.megastudy.net", "mgangdong.megastudy.net", "tgangdong.megastudy.net", "tmgangdong.megastudy.net"
					
					gMGC_Campus_N = "gangdong"
					gMGC_Campus_Code = "CD0198"
					gMGC_Campus_KN = "���� �ް����͵��п�"
					gMGC_Campus_EN = "gangdong"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "�����"
					gMGC_Campus_Detail_Code = "CD0240"
					gMGC_Campus_DirSite ="/gangdong/nsu"
					gMGC_Campus_Title = cDev&"���� �ް����͵��п�"
					gMGC_Campus_Theme = "red"

				'���� ��.��.��	http://gangdongs.megastudy.net	http://mgangdongs.megastudy.net
				Case "gangdongs.megastudy.net", "mgangdongs.megastudy.net", "tgangdongs.megastudy.net", "tmgangdongs.megastudy.net"
					
					gMGC_Campus_N = "gangdong"
					gMGC_Campus_Code = "CD0198"
					gMGC_Campus_KN = "���� �ް����͵��п�"
					gMGC_Campus_EN = "gangdong"
					gMGC_Campus_Kind = "jaehak"
					gMGC_Campus_Kind_KN = "���߰��"
					gMGC_Campus_Detail_Code = "CD0215"
					gMGC_Campus_DirSite ="/gangdong/jaehak"
					gMGC_Campus_Title = cDev&"���� �ް����͵��п�"
					gMGC_Campus_Theme = "blue" 
				
				'���� �����	http://gangbuk.megastudy.net	http://mgangbuk.megastudy.net
				Case "gangbuk.megastudy.net", "mgangbuk.megastudy.net", "tgangbuk.megastudy.net", "tmgangbuk.megastudy.net"
				
					gMGC_Campus_N = "gangbuk"
					gMGC_Campus_Code = "CD0005"
					gMGC_Campus_KN = "���� �ް����͵��п�"
					gMGC_Campus_EN = "gangbuk"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "�����"
					gMGC_Campus_Detail_Code = "CD0210"
					gMGC_Campus_DirSite ="/gangbuk/nsu"
					gMGC_Campus_Title = cDev&"���� �ް����͵��п�"
					gMGC_Campus_Theme = "red" 

				'���� ����	http://gangbuks.megastudy.net	http://mgangbuks.megastudy.net
				Case "gangbuks.megastudy.net", "mgangbuks.megastudy.net", "tgangbuks.megastudy.net", "tmgangbuks.megastudy.net"
				
					gMGC_Campus_N = "gangbuk"
					gMGC_Campus_Code = "CD0005"
					gMGC_Campus_KN = "���� �ް����͵��п�"
					gMGC_Campus_EN = "gangbuk"
					gMGC_Campus_Kind = "jaehak"
					gMGC_Campus_Kind_KN = "��� ���ա��ܰ��п�"
					gMGC_Campus_Detail_Code = "CD0209"
					gMGC_Campus_DirSite ="/gangbuk/jaehak"
					gMGC_Campus_Title = cDev&"���� �ް����͵��п� ����"
					gMGC_Campus_Theme = "blue"  

				'�뷮�� �����	http://noryangjin.megastudy.net	http://mnoryangjin.megastudy.net
				Case "noryangjin.megastudy.net", "mnoryangjin.megastudy.net", "tnoryangjin.megastudy.net", "tmnoryangjin.megastudy.net"
					
					gMGC_Campus_N = "noryangjin"
					gMGC_Campus_Code = "CD0006"
					gMGC_Campus_KN = "�뷮�� �ް����͵��п�"
					gMGC_Campus_EN = "noryangjin"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "�����"
					gMGC_Campus_Detail_Code = "CD0211"
					gMGC_Campus_DirSite ="/noryangjin/nsu"
					gMGC_Campus_Title = cDev&"�뷮�� �ް����͵��п�"
					gMGC_Campus_Theme = "red"   

				'�뷮�� �ܰ�	http://nrj.megastudy.net	http://mnrj.megastudy.net
				Case "nrj.megastudy.net", "mnrj.megastudy.net","tnrj.megastudy.net", "tmnrj.megastudy.net"
					
					gMGC_Campus_N = "noryangjin"
					gMGC_Campus_Code = "CD0006"
					gMGC_Campus_KN = "�뷮�� �ް����͵��п�"
					gMGC_Campus_EN = "noryangjin"
					gMGC_Campus_Kind = "danka"
					gMGC_Campus_Kind_KN = "�ܰ�"
					gMGC_Campus_Detail_Code = "CD0212"
					gMGC_Campus_DirSite ="/noryangjin/danka"
					gMGC_Campus_Title = cDev&"�뷮�� �ް����͵��п� �ܰ�"
					gMGC_Campus_Theme = "blue"   
				
				'���� �����	http://seocho.megastudy.net	http://mseocho.megastudy.net
				Case "seocho.megastudy.net", "mseocho.megastudy.net", "tseocho.megastudy.net", "tmseocho.megastudy.net"
					
					gMGC_Campus_N = "seocho"
					gMGC_Campus_Code = "CD0004"
					gMGC_Campus_KN = "���� �ް����͵��п�"
					gMGC_Campus_EN = "seocho"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "�����"
					gMGC_Campus_Detail_Code = "CD0208"
					gMGC_Campus_DirSite ="/seocho/nsu"
					gMGC_Campus_Title = cDev&"���� �ް����͵��п�"
					gMGC_Campus_Theme = "red" 

				'���� ����	http://seochos.megastudy.net	http://mseochos.megastudy.net
				Case "seochos.megastudy.net", "mseochos.megastudy.net", "tseochos.megastudy.net", "tmseochos.megastudy.net"
					
					gMGC_Campus_N = "seocho"
					gMGC_Campus_Code = "CD0004"
					gMGC_Campus_KN = "���� �ް����͵��п�"
					gMGC_Campus_EN = "seocho"
					gMGC_Campus_Kind = "jaehak"
					gMGC_Campus_Kind_KN = "����"
					gMGC_Campus_Detail_Code = "CD0207"
					gMGC_Campus_DirSite ="/seocho/jaehak"
					gMGC_Campus_Title = cDev&"���� �ް����͵��п� ����"
					gMGC_Campus_Theme = "blue"   
					
				'���� �����	http://seongbuk.megastudy.net	http://mseongbuk.megastudy.net
				Case "seongbuk.megastudy.net", "mseongbuk.megastudy.net", "tseongbuk.megastudy.net", "tmseongbuk.megastudy.net"
					
					gMGC_Campus_N = "seongbuk"
					gMGC_Campus_Code = "CD0098"
					gMGC_Campus_KN = "���� �ް����͵��п�"
					gMGC_Campus_EN = "seongbuk"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "�����"
					gMGC_Campus_Detail_Code = "CD0239"
					gMGC_Campus_DirSite ="/seongbuk/nsu"
					gMGC_Campus_Title = cDev&"���� �ް����͵��п�"
					gMGC_Campus_Theme = "red"  
					
				'���� ��.��.��	http://seongbuks.megastudy.net	 http://mseongbuks.megastudy.net
				Case "seongbuks.megastudy.net", "mseongbuks.megastudy.net", "tseongbuks.megastudy.net", "tmseongbuks.megastudy.net"
					
					gMGC_Campus_N = "seongbuk"
					gMGC_Campus_Code = "CD0098"
					gMGC_Campus_KN = "���� �ް����͵��п�"
					gMGC_Campus_EN = "seongbuk"
					gMGC_Campus_Kind = "jaehak"
					gMGC_Campus_Kind_KN = "���߰��"
					gMGC_Campus_Detail_Code = "CD0214"
					gMGC_Campus_DirSite ="/seongbuk/jaehak"
					gMGC_Campus_Title = cDev&"���� �ް����͵��п�"
					gMGC_Campus_Theme = "blue"  
				
				'���� ������Թ�	http://sinchon.megastudy.net	http://msinchon.megastudy.net
				Case "sinchon.megastudy.net", "msinchon.megastudy.net", "tsinchon.megastudy.net", "tmsinchon.megastudy.net"
					
					gMGC_Campus_N = "sinchon"
					gMGC_Campus_Code = "CD0036"
					gMGC_Campus_KN = "���� �ް����͵��п�"
					gMGC_Campus_EN = "sinchon"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "���/�ܰ�"
					gMGC_Campus_Detail_Code = "CD0213"
					gMGC_Campus_DirSite ="/sinchon"
					gMGC_Campus_Title = cDev&"���� �ް����͵��п�"
					gMGC_Campus_Theme = "red"  

				'���� �����	http://pyeongchon.megastudy.net	http://mpyeongchon.megastudy.net
				Case "pyeongchon.megastudy.net", "mpyeongchon.megastudy.net" , "tpyeongchon.megastudy.net", "tmpyeongchon.megastudy.net"
					
					gMGC_Campus_N = "pyeongchon"
					gMGC_Campus_Code = "CD0179"
					gMGC_Campus_KN = "���� �ް����͵��п�"
					gMGC_Campus_EN = "pyeongchon"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "�����"
					gMGC_Campus_Detail_Code = "CD0217"
					gMGC_Campus_DirSite ="/pyeongchon/nsu"
					gMGC_Campus_Title = cDev&"���� �ް����͵��п�"
					gMGC_Campus_Theme = "red"  

				'���� ����	http://pyeongchons.megastudy.net	http://mpyeongchons.megastudy.net
				Case "pyeongchons.megastudy.net", "mpyeongchons.megastudy.net","tpyeongchons.megastudy.net", "tmpyeongchons.megastudy.net"
					
					gMGC_Campus_N = "pyeongchon"
					gMGC_Campus_Code = "CD0179"
					gMGC_Campus_KN = "���� �ް����͵��п�"
					gMGC_Campus_EN = "pyeongchon"
					gMGC_Campus_Kind = "jaehak"
					gMGC_Campus_Kind_KN = "����"
					gMGC_Campus_Detail_Code = "CD0216"
					gMGC_Campus_DirSite ="/pyeongchon/jaehak"
					gMGC_Campus_Title = cDev&"���� �ް����͵��п� ����"
					gMGC_Campus_Theme = "blue"  

				'���� ���� ������Թ�	http://yangji.megastudy.net	http://myangji.megastudy.net
				Case "yangji.megastudy.net", "myangji.megastudy.net", "tyangji.megastudy.net", "tmyangji.megastudy.net"	
					
					gMGC_Campus_N = "yangji"
					gMGC_Campus_Code = "CD0178"
					gMGC_Campus_KN = "�ֻ����� ������ ���� ����п�"
					gMGC_Campus_EN = "yangji"
					gMGC_Campus_Kind = "gisuk"
					gMGC_Campus_Kind_KN = ""
					gMGC_Campus_Detail_Code = "CD0220"
					gMGC_Campus_DirSite ="/yangji"
					gMGC_Campus_Title = cDev&"�ֻ����� ������ ���� ����п�"
					gMGC_Campus_Theme = "beige" 

				'���� �Ű� ������Թ�	http://yangjim.megastudy.net	http://myangjim.megastudy.net
				Case "yangjim.megastudy.net", "myangjim.megastudy.net"	, "tyangjim.megastudy.net", "tmyangjim.megastudy.net" ,"seochob.megastudy.net", "mseochob.megastudy.net"	, "tseochob.megastudy.net", "tmseochob.megastudy.net"
					
					gMGC_Campus_N = "yangjim"
					gMGC_Campus_Code = "CD0039"
					gMGC_Campus_KN = "�ڿ��� ������ ���� ����п�"
					gMGC_Campus_EN = "yangjim"
					gMGC_Campus_Kind = "gisuk"
					gMGC_Campus_Kind_KN = ""
					gMGC_Campus_Detail_Code = "CD0243"
					gMGC_Campus_DirSite ="/yangjim"
					gMGC_Campus_Title = cDev&"�ڿ��� ������ ���� ����п�"
					gMGC_Campus_Theme = "beige"   'purple
				

				'������ġ http://campus.magastudy.net/russel/
				Case "russeldc.megastudy.net", "mrusseldc.megastudy.net"

					 gMGC_Campus_N = "russel"
					 gMGC_Campus_Code = "CD0001"
					 gMGC_Campus_KN = "�ް����͵� ���� ��ġ"
					 gMGC_Campus_EN = "russel"
					 gMGC_Campus_Kind = "danka"
					 gMGC_Campus_Kind_KN = "�ܰ�"
					 gMGC_Campus_Detail_Code = "CD0241"
					 If gMGC_RequestServerName ="mrusseldc.megastudy.net" Then 
					 gMGC_Campus_DirSite ="/m_russel/"
					 Else 
					 gMGC_Campus_DirSite ="/russel/"
					 End If 
					 gMGC_Campus_Title = cDev&"���� ��ġ "
					 gMGC_Campus_Theme = ""

			   '�����д� http://campus.magastudy.net/russel_bundang/
			   Case "russelbd.megastudy.net", "mrusselbd.megastudy.net"

					gMGC_Campus_N = "russel_bundang"
					gMGC_Campus_Code = "CD0244"
					gMGC_Campus_KN = "�ް����͵� ���� �д�"
					gMGC_Campus_EN = "russel_bundang"
					gMGC_Campus_Kind = "danka"
					gMGC_Campus_Kind_KN = "�ܰ�"
					gMGC_Campus_Detail_Code = "CD0244"
					gMGC_Campus_DirSite ="/russel_bundang"
					gMGC_Campus_Title = cDev&"���� �д�"
					gMGC_Campus_Theme = ""

			   '������ http://campus.magastudy.net/russel_mokdong/
				Case "russelmd.megastudy.net", "mrusselmd.megastudy.net"

					gMGC_Campus_N = "russel_mokdong"
					gMGC_Campus_Code = "CD0245"
					gMGC_Campus_KN = "�ް����͵� ���� ��"
					gMGC_Campus_EN = "russel_mokdong"
					gMGC_Campus_Kind = "danka"
					gMGC_Campus_Kind_KN = "�ܰ�"
					gMGC_Campus_Detail_Code = "CD0245"
					gMGC_Campus_DirSite ="/russel_mokdong"
					gMGC_Campus_Title = cDev&"���� ��"
					gMGC_Campus_Theme = ""
				
				Case "russelct.megastudy.net", "russelct.megastudy.net"

					gMGC_Campus_N = "russel_ct"
					gMGC_Campus_Code = "CD0246"
					gMGC_Campus_KN = "�ް����͵� ���� ����"
					gMGC_Campus_EN = "russel_ct"
					gMGC_Campus_Kind = "danka"
					gMGC_Campus_Kind_KN = "�ܰ�"
					gMGC_Campus_Detail_Code = "CD0246"
					gMGC_Campus_DirSite ="/russel_ct"
					gMGC_Campus_Title = cDev&"���� ����"
					gMGC_Campus_Theme = ""

				Case "russelgn.megastudy.net", "russelgn.megastudy.net"

					gMGC_Campus_N = "russel_gn"
					gMGC_Campus_Code = "CD0247"
					gMGC_Campus_KN = "�ް����͵� ���� ����"
					gMGC_Campus_EN = "russel_gn"
					gMGC_Campus_Kind = "danka"
					gMGC_Campus_Kind_KN = "�ܰ�"
					gMGC_Campus_Detail_Code = "CD0247"
					gMGC_Campus_DirSite ="/russel_gn"
					gMGC_Campus_Title = cDev&"���� ����"
					gMGC_Campus_Theme = ""
				
				Case "russelpc.megastudy.net", "russelpc.megastudy.net"

					gMGC_Campus_N = "russel_pc"
					gMGC_Campus_Code = "CD0248"
					gMGC_Campus_KN = "�ް����͵� ���� ����"
					gMGC_Campus_EN = "russel_pc"
					gMGC_Campus_Kind = "danka"
					gMGC_Campus_Kind_KN = "�ܰ�"
					gMGC_Campus_Detail_Code = "CD0248"
					gMGC_Campus_DirSite ="/russel_pc"
					gMGC_Campus_Title = cDev&"���� ����"
					gMGC_Campus_Theme = ""

				
				Case "russelyt.megastudy.net", "russelyt.megastudy.net"

					gMGC_Campus_N = "russel_yt"
					gMGC_Campus_Code = "CD0249"
					gMGC_Campus_KN = "�ް����͵� ���� ����"
					gMGC_Campus_EN = "russel_yt"
					gMGC_Campus_Kind = "danka"
					gMGC_Campus_Kind_KN = "�ܰ�"
					gMGC_Campus_Detail_Code = "CD0249"
					gMGC_Campus_DirSite ="/russel_yt"
					gMGC_Campus_Title = cDev&"���� ����"
					gMGC_Campus_Theme = ""

				
				Case "russelbc.megastudy.net", "russelbc.megastudy.net"

					gMGC_Campus_N = "russel_bc"
					gMGC_Campus_Code = "CD0250"
					gMGC_Campus_KN = "�ް����͵� ���� ��õ"
					gMGC_Campus_EN = "russel_bc"
					gMGC_Campus_Kind = "danka"
					gMGC_Campus_Kind_KN = "�ܰ�"
					gMGC_Campus_Detail_Code = "CD0250"
					gMGC_Campus_DirSite ="/russel_bc"
					gMGC_Campus_Title = cDev&"���� ��õ"
					gMGC_Campus_Theme = ""
				
				Case "russel.megastudy.net"

					gMGC_Campus_N = "intro"
					gMGC_Campus_Code = "RUSSELINTRO"
					gMGC_Campus_KN = "�ް����͵� ����"
					gMGC_Campus_EN = "intro"
					gMGC_Campus_Kind = "danka"
					gMGC_Campus_Kind_KN = "�ܰ�"
					gMGC_Campus_Detail_Code = "RUSSELINTRO"
					gMGC_Campus_DirSite ="/intro"
					gMGC_Campus_Title = cDev&"���� INTRO"
					gMGC_Campus_Theme = ""
				
				'�����߰� http://campus.magastudy.net/russel_jg/
				Case "russeljg.megastudy.net", "mrusseljg.megastudy.net"
					gMGC_Campus_N = "russel_jg"
					gMGC_Campus_Code = "CD0257"
					gMGC_Campus_KN = "�ް����͵� ���� �߰�"
					gMGC_Campus_EN = "russel_jg"
					gMGC_Campus_Kind = "danka"
					gMGC_Campus_Kind_KN = "�ܰ�"
					gMGC_Campus_Detail_Code = "CD0257"
					gMGC_Campus_DirSite ="/russel_jg/"
					gMGC_Campus_Title = cDev&"���� �߰�"
					gMGC_Campus_Theme = ""

				'���� ���
				Case "russelcampus.megastudy.net", "mrusselcampus.megastudy.net"
					gMGC_Campus_N = "russel_yj"
					gMGC_Campus_Code = "CD0258"
					gMGC_Campus_KN = "�ް����͵� ���� ���"
					gMGC_Campus_EN = "russel_yj"
					gMGC_Campus_Kind = "gisuk"
					gMGC_Campus_Kind_KN = "���"
					gMGC_Campus_Detail_Code = "CD0258"
					gMGC_Campus_DirSite ="/russel_yj/"
					gMGC_Campus_Title = cDev&"���� ���"
					gMGC_Campus_Theme = ""


				'��õ ������Թ�	http://bucheon.megastudy.net	http://mbucheon.megastudy.net
				Case "bucheon.megastudy.net", "mbucheon.megastudy.net", "tbucheon.megastudy.net", "tmbucheon.megastudy.net"
				
					gMGC_Campus_N = "bucheon"
					gMGC_Campus_Code = "CD0250"
					gMGC_Campus_KN = "��õ �ް����͵��п�"
					gMGC_Campus_EN = "bucheon"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "�����"
					gMGC_Campus_Detail_Code = "CD0251"
					gMGC_Campus_DirSite ="/bucheon/nsu"
					gMGC_Campus_Title = cDev&"��õ �ް����͵��п�"
					gMGC_Campus_Theme = "red" 

				'�д� ������Թ�	http://bundang.megastudy.net	http://mbundang.megastudy.net
				Case "bundang.megastudy.net", "mbundang.megastudy.net", "tbundang.megastudy.net", "tmbundang.megastudy.net"
				
					gMGC_Campus_N = "bundang"
					gMGC_Campus_Code = "CD0252"
					gMGC_Campus_KN = "�д� �ް����͵��п�"
					gMGC_Campus_EN = "bundang"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "�����"
					gMGC_Campus_Detail_Code = "CD0253"
					gMGC_Campus_DirSite ="/bundang/nsu"
					gMGC_Campus_Title = cDev&"�д� �ް����͵��п�"
					gMGC_Campus_Theme = "red" 
			

					'�ϻ� ������Թ�	http://ilsan.megastudy.net	http://milsan.megastudy.net
				Case "ilsan.megastudy.net", "milsan.megastudy.net", "tilsan.megastudy.net", "tmilsan.megastudy.net"
				
					gMGC_Campus_N = "ilsan"
					gMGC_Campus_Code = "CD0254"
					gMGC_Campus_KN = "�ϻ� �ް����͵��п�"
					gMGC_Campus_EN = "ilsan"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "�����"
					gMGC_Campus_Detail_Code = "CD0255"
					gMGC_Campus_DirSite ="/ilsan/nsu"
					gMGC_Campus_Title = cDev&"�ϻ� �ް����͵��п�"
					gMGC_Campus_Theme = "red" 

				'���� ��.��.��	http://yangju.megastudy.net	 http://myangju.megastudy.net
				Case "yangju.megastudy.net", "myangju.megastudy.net", "tyangju.megastudy.net", "tmyangju.megastudy.net"

					gMGC_Campus_N = "yangju"
					gMGC_Campus_Code = "CD0258"
					gMGC_Campus_KN = "���� �ް����͵��п�"
					gMGC_Campus_EN = "yangju"
					gMGC_Campus_Kind = "jaehak"
					gMGC_Campus_Kind_KN = "���߰��"
					gMGC_Campus_Detail_Code = "CD0259"
					gMGC_Campus_DirSite ="/yangju/jaehak"
					gMGC_Campus_Title = cDev&"���� �ް����͵��п�"
					gMGC_Campus_Theme = "blue"

				'���� ��.��.��	http://unjung.megastudy.net	 http://munjung.megastudy.net
				Case "unjung.megastudy.net", "munjung.megastudy.net", "tunjung.megastudy.net", "tmunjung.megastudy.net"

					gMGC_Campus_N = "unjung"
					gMGC_Campus_Code = "CD0260"
					gMGC_Campus_KN = "���� �ް����͵��п�"
					gMGC_Campus_EN = "unjung"
					gMGC_Campus_Kind = "jaehak"
					gMGC_Campus_Kind_KN = "�߰���"
					gMGC_Campus_Detail_Code = "CD0261"
					gMGC_Campus_DirSite ="/unjung/jaehak"
					gMGC_Campus_Title = cDev&"���� �ް����͵��п�"
					gMGC_Campus_Theme = "blue"

				'���� �����	http://songpa.megastudy.net	http://msongpa.megastudy.net
				Case "songpa.megastudy.net", "msongpa.megastudy.net", "tsongpa.megastudy.net", "tmsongpa.megastudy.net"
					
					gMGC_Campus_N = "songpa"
					gMGC_Campus_Code = "CD0275"
					gMGC_Campus_KN = "���� �ް����͵��п�"
					gMGC_Campus_EN = "songpa"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "�����"
					gMGC_Campus_Detail_Code = "CD0276"
					gMGC_Campus_DirSite ="/songpa/nsu"
					gMGC_Campus_Title = cDev&"���� �ް����͵��п�"
					gMGC_Campus_Theme = "red"

				'�ȼ���� ������Թ�	http://ansung.megastudy.net	http://mansung.megastudy.net
				Case "ansung.megastudy.net", "mansung.megastudy.net", "tansung.megastudy.net", "tmansung.megastudy.net"
					
					gMGC_Campus_N = "ansung"
					gMGC_Campus_Code = "CD0279"
					gMGC_Campus_KN = "�ȼ� �ް����͵� ����п�"
					gMGC_Campus_EN = "ansung"
					gMGC_Campus_Kind = "gisuk"
					gMGC_Campus_Kind_KN = ""
					gMGC_Campus_Detail_Code = "CD0280"
					gMGC_Campus_DirSite ="/ansung"
					gMGC_Campus_Title = cDev&"�ȼ� �ް����͵� ����п�"
					gMGC_Campus_Theme = "beige"   'purple

				Case Else

					gMGC_Campus_N = "campus"
					gMGC_Campus_Code = ""
					gMGC_Campus_KN = ""
					gMGC_Campus_EN = ""
					gMGC_Campus_Kind = ""
					gMGC_Campus_Kind_KN = ""
					gMGC_Campus_Detail_Code = ""
					gMGC_Campus_DirSite ="/campus"	
					gMGC_Campus_Title = cDev&"�ް����͵��п�"
					gMGC_Campus_Theme = ""    
				
			End Select

		End If	' ���� �۾� ���а� ��

		'����Ʈ ���� ���� E

		returnVal = ""
		
		'�ҽ����̷� ���� �����߻� ���� �ϱ� ���� �ӽ� ���� 
		If ValueCheck(Request.Cookies("_CAMPUS_CODE_") ) Then 
			TmpCheckValue = True 
		Else 
			TmpCheckValue = False  
		End If 

		TmpCheckValue = False  

		If TmpCheckValue Then 
			'������ ��Ű �� ���� S
			Select Case UCase(CStr(VarsName))
				Case "_CAMPUS_" '�п� ����
					returnVal = Request.Cookies("_CAMPUS_")
				Case "_CAMPUS_CODE_" '�п� ���ڵ� ����
					returnVal = Request.Cookies("_CAMPUS_CODE_")
				Case "_CAMPUS_NAME_" '�п��̸� ����
					returnVal = Request.Cookies("_CAMPUS_NAME_")
				Case "_CAMPUS_ENG_NAME_" '�п� ������
					returnVal = Request.Cookies("_CAMPUS_ENG_NAME_")
				Case "_CAMPUS_KIND_" '�п� �ܰ�,���,��� ����
					returnVal = Request.Cookies("_CAMPUS_KIND_")
					'����� returnVal = Request.Cookies("_CAMPUS_CATE_")
				Case "_CAMPUS_DETAIL_CODE_" '�п� ���ڵ� ����			
					returnVal = Request.Cookies("_CAMPUS_DETAIL_CODE_")
					'����� returnVal = Request.Cookies("_CAMPUS_CODE_")
				Case "_CAMPUS_DIRSITE_" '�п� ���丮 ����			
					returnVal = Request.Cookies("_CAMPUS_DIRSITE_")
				Case "_CAMPUS_KIND_NAME_" '�п� ī�װ� �̸� ���� 
					returnVal = Request.Cookies("_CAMPUS_CATE_NAME_")
				Case "_CAMPUS_THEME_" '����� �п� �׸� 
					returnVal = Request.Cookies("_CAMPUS_THEME_")
				Case Else '�׿� ����
					returnVal = ""
			End Select
			'������ ��Ű �� ���� E
		Else 
			'������ ���� �� ���� S
			Select Case UCase(CStr(VarsName))
				Case "_CAMPUS_" '�п� ����
					returnVal = gMGC_Campus_N
				Case "_CAMPUS_CODE_" '�п� ���ڵ� ����
					returnVal = gMGC_Campus_Code
				Case "_CAMPUS_NAME_" '�п��̸� ����
					returnVal = gMGC_Campus_KN
				Case "_CAMPUS_ENG_NAME_" '�п� ������
					returnVal = gMGC_Campus_EN	
				Case "_CAMPUS_KIND_" '�п� �ܰ�,���,��� ����
					returnVal = gMGC_Campus_Kind
				Case "_CAMPUS_DETAIL_CODE_" '�п� ���ڵ� ����			
					returnVal = gMGC_Campus_Detail_Code
				Case "_CAMPUS_DIRSITE_" '�п� ���丮 ����			
					returnVal = gMGC_Campus_DirSite
				Case "_CAMPUS_KIND_NAME_" '�п� ī�װ� �̸� ���� 
					returnVal = gMGC_Campus_Kind_KN	
				Case "_CAMPUS_THEME_" '����� �п� �׸� 
					returnVal = gMGC_Campus_Theme
				Case "_CAMPUS_TITLE_" '����� �п� �׸� 
					returnVal = gMGC_Campus_Title
				Case Else '�׿� ����
					returnVal = ""
			End Select
			'������ ���� �� ���� E
		End If 

		GetCampusVarsSiteMGC =  returnVal
	End Function 
	'�п��� ���� �������� �Լ� E

	' �п��� ����Ʈ�ּ� ��������  S
	Function GetCampusDetailUrlSiteMGC(sCampusCodeDetail,isDirCheck)
		returnURLSite = ""
		isCheckDev = False 

		If Request.ServerVariables("LOCAL_ADDR") = "10.1.3.10" Then  '���߼��� üũ
			isCheckDev = True 
		Else
			isCheckDev = False 
		End If

		If Request.ServerVariables("LOCAL_ADDR") = "::1" Then	' ���ÿ��� �۾��� ���ؼ� �߰� (2020.03.10)
			Select Case UCase(CStr(sCampusCodeDetail))
				Case "INTRO" '��Ʈ����������
					If isDirCheck Then 
						returnURLSite = "/campus"
					End If 
				Case "CD0205" '��������
					If isDirCheck Then 
						returnURLSite = "/gangnam/jaehak"
					End If 
				Case "CD0206" '���������
					If isDirCheck Then 
						returnURLSite = "/gangnam/nsu"
					End If 
				Case "CD0207" '���ʰ���
					If isDirCheck Then 
						returnURLSite = "/seocho/jaehak"
					End If 
				Case "CD0208" '���������
					If isDirCheck Then 
						returnURLSite = "/seocho/nsu"
					End If 
				Case "CD0209" '���ϰ���
					If isDirCheck Then 
						returnURLSite = "/gangbuk/jaehak"
					End If 
				Case "CD0210" '���������
					If isDirCheck Then 
						returnURLSite = "/gangbuk/nsu"
					End If 
				Case "CD0211" '�뷮�������
					If isDirCheck Then 
						returnURLSite = "/noryangjin/nsu"
					End If 
				Case "CD0212" '�뷮���ܰ�
					If isDirCheck Then 
						returnURLSite = "/noryangjin/danka"
					End If 
				Case "CD0213" '����
					If isDirCheck Then 
						returnURLSite = "/sinchon"
					End If 
				Case "CD0214" '���ϰ���
					If isDirCheck Then 
						returnURLSite = "/seongbuk/jaehak"
					End If 
				Case "CD0239" '���������
					If isDirCheck Then 
						returnURLSite = "/seongbuk/nsu"
					End If 
				Case "CD0215" '��������
					If isDirCheck Then 
						returnURLSite = "/gangdong/jaehak"
					End If 

				Case "CD0240" '���������
					If isDirCheck Then 
						returnURLSite = "/gangdong/nsu"
					End If 
				Case "CD0216" '���̰���
					If isDirCheck Then 
						returnURLSite = "/pyeongchon/jaehak"
					End If 
				Case "CD0217" '���������
					If isDirCheck Then 
						returnURLSite = "/pyeongchon/nsu"
					End If 
				Case "CD0218" '����
					If isDirCheck Then 
						returnURLSite = returnURLSite & ""
					End If 
				Case "CD0219" '������
					If isCheckDev Then 
						returnURLSite = ""
					Else 
						returnURLSite = ""
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & ""
					End If 
				Case "CD0220" '��������
					If isDirCheck Then 
						returnURLSite = "/yangji"
					End If 
				Case "CD0243" '�����Ű�
					If isDirCheck Then 
						returnURLSite = "/yangjim"
					End If 
				Case "CD0221" '���� -- ���Ͼ���
					If isCheckDev Then 
						returnURLSite = ""
					Else 
						returnURLSite = ""
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & ""
					End If 
				Case "CD0241" '���� ��ġ
					  If isDirCheck Then 
						returnURLSite = "/russel"
					  End If 

				Case "CD0244" '���� �д�
					  If isDirCheck Then 
						 returnURLSite = "/russel_bundang"
					  End If 

				Case "CD0245" '���� ��
						If isDirCheck Then 
						  returnURLSite = "/russel_mokdong"
						End If
				
				 Case "CD0246" '���� ����
						If isDirCheck Then 
						  returnURLSite = "/russel_ct"
						End If

				Case "CD0247" '���� ����
						If isDirCheck Then 
						  returnURLSite = "/russel_gn"
						End If
				
				Case "CD0248" '���� ����
						If isDirCheck Then 
						  returnURLSite = "/russel_pc"
						End If
				
				Case "CD0249" '���� ����
						If isDirCheck Then 
						  returnURLSite = "/russel_yt"
						End If

				Case "CD0250" '���� ��õ
						If isDirCheck Then 
						  returnURLSite = "/russel_bc"
						End If
				Case "RUSSELINTRO" '���� INTRO
						If isDirCheck Then 
						  returnURLSite = "/intro"
						End If


				Case "CD0251" '��õ �����
					If isDirCheck Then 
						returnURLSite = "/bucheon/nsu"
					End If 

				Case "CD0253" '�д� �����
					If isDirCheck Then 
						returnURLSite = "/bundang/nsu"
					End If 


				Case "CD0255" '�ϻ� �����
					If isDirCheck Then 
						returnURLSite = "/ilsan/nsu"
					End If 


				Case "CD0259" '���ְ���
					If isDirCheck Then
						returnURLSite = "/yangju/jaehak"
					End If

				Case "CD0261" '��������
					If isDirCheck Then
						returnURLSite = "/unjung/jaehak"
					End If

				Case "CD0276" '���������
					If isDirCheck Then 
						returnURLSite = "/songpa/nsu"
					End If 

				Case "CD0280" '�ȼ����
					If isDirCheck Then 
						returnURLSite = "/ansung"
					End If 

				Case Else '�� �̿��ǰ�� 
					If isDirCheck Then 
						returnURLSite = "/campus"
					End If 

			End Select

		Else

			Select Case UCase(CStr(sCampusCodeDetail))
				Case "INTRO" '��Ʈ����������
					If isCheckDev Then 
						returnURLSite = "http://tcampus.megastudy.net"
					Else 
						returnURLSite = "http://campus.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/campus"
					End If 
				Case "CD0205" '��������
					If isCheckDev Then 
						returnURLSite = "http://tgangnam.megastudy.net"
					Else 
						returnURLSite = "http://gangnam.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/gangnam/jaehak"
					End If 
				Case "CD0206" '���������
					If isCheckDev Then 
						returnURLSite = "http://tgangnam.megastudy.net"
					Else 
						returnURLSite = "http://gangnam.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/gangnam/nsu"
					End If 
				Case "CD0207" '���ʰ���
					If isCheckDev Then 
						returnURLSite = "http://tseocho.megastudy.net"
					Else 
						returnURLSite = "http://seocho.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/seocho/jaehak"
					End If 
				Case "CD0208" '���������
					If isCheckDev Then 
						returnURLSite = "http://tseocho.megastudy.net"
					Else 
						returnURLSite = "http://seocho.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/seocho/nsu"
					End If 
				Case "CD0209" '���ϰ���
					If isCheckDev Then 
						returnURLSite = "http://tgangbuks.megastudy.net"
					Else 
						returnURLSite = "http://gangbuks.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/gangbuk/jaehak"
					End If 
				Case "CD0210" '���������
					If isCheckDev Then 
						returnURLSite = "http://tgangbuk.megastudy.net"
					Else 
						returnURLSite = "http://gangbuk.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/gangbuk/nsu"
					End If 
				Case "CD0211" '�뷮�������
					If isCheckDev Then 
						returnURLSite = "http://tnoryangjin.megastudy.net"
					Else 
						returnURLSite = "http://noryangjin.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/noryangjin/nsu"
					End If 
				Case "CD0212" '�뷮���ܰ�
					If isCheckDev Then 
						returnURLSite = "http://tnoryangjin.megastudy.net"
					Else 
						returnURLSite = "http://noryangjin.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/noryangjin/danka"
					End If 
				Case "CD0213" '����
					If isCheckDev Then 
						returnURLSite = "http://tsinchon.megastudy.net"
					Else 
						returnURLSite = "http://sinchon.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/sinchon"
					End If 
				Case "CD0214" '���ϰ���
					If isCheckDev Then 
						returnURLSite = "http://tseongbuks.megastudy.net"
					Else 
						returnURLSite = "http://seongbuks.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/seongbuk/jaehak"
					End If 
				Case "CD0239" '���������
					If isCheckDev Then 
						returnURLSite = "http://tseongbuk.megastudy.net"
					Else 
						returnURLSite = "http://seongbuk.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/seongbuk/nsu"
					End If 
				Case "CD0215" '��������
					If isCheckDev Then 
						returnURLSite = "http://tgangdongs.megastudy.net"
					Else 
						returnURLSite = "http://gangdongs.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/gangdong/jaehak"
					End If 

				Case "CD0240" '���������
					If isCheckDev Then 
						returnURLSite = "http://tgangdong.megastudy.net"
					Else 
						returnURLSite = "http://gangdong.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/gangdong/nsu"
					End If 
				Case "CD0216" '���̰���
					If isCheckDev Then 
						returnURLSite = "http://tpyeongchon.megastudy.net"
					Else 
						returnURLSite = "http://pyeongchon.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/pyeongchon/jaehak"
					End If 
				Case "CD0217" '���������
					If isCheckDev Then 
						returnURLSite = "http://tpyeongchon.megastudy.net"
					Else 
						returnURLSite = "http://pyeongchon.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/pyeongchon/nsu"
					End If 
				Case "CD0218" '����
					If isCheckDev Then 
						returnURLSite = ""
					Else 
						returnURLSite = ""
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & ""
					End If 
				Case "CD0219" '������
					If isCheckDev Then 
						returnURLSite = ""
					Else 
						returnURLSite = ""
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & ""
					End If 
				Case "CD0220" '��������
					If isCheckDev Then 
						returnURLSite = "http://tyangji.megastudy.net"
					Else 
						returnURLSite = "http://yangji.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/yangji"
					End If 
				Case "CD0243" '�����Ű�
					If isCheckDev Then 
						'returnURLSite = "http://tyangjim.megastudy.net"
						returnURLSite = "http://tseochob.megastudy.net"
					Else 
						'returnURLSite = "http://yangjim.megastudy.net"
						returnURLSite = "http://seochob.megastudy.net"
					End If 

					gMGC_RequestServerName_SUB = LCase(Request.ServerVariables("SERVER_NAME"))  '������ �̸�

					If gMGC_RequestServerName_SUB = "seochob.megastudy.net" Or gMGC_RequestServerName_SUB = "mseochob.megastudy.net" Or gMGC_RequestServerName_SUB ="tseochob.megastudy.net" Or gMGC_RequestServerName_SUB = "tmseochob.megastudy.net" Then 
						returnURLSite = "http://"&gMGC_RequestServerName_SUB
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/yangjim"
					End If 
				Case "CD0221" '���� -- ���Ͼ���
					If isCheckDev Then 
						returnURLSite = ""
					Else 
						returnURLSite = ""
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & ""
					End If 
				Case "CD0241" '���� ��ġ
					  If isCheckDev Then 
						 returnURLSite = "http://russeldc.megastudy.net"
					  Else 
						 returnURLSite = "http://russeldc.megastudy.net"
					  End If 

					  If isDirCheck Then 
						returnURLSite = returnURLSite & "/russel"
					  End If 

				Case "CD0244" '���� �д�
					  If isCheckDev Then 
						 returnURLSite = "http://russelbd.megastudy.net"
					  Else 
						 returnURLSite = "http://russelbd.megastudy.net"
					  End If 

					  If isDirCheck Then 
						 returnURLSite = returnURLSite & "/russel_bundang"
					  End If 

				Case "CD0245" '���� ��
					   If isCheckDev Then 
						 returnURLSite = "http://russelmd.megastudy.net"
					   Else 
						 returnURLSite = "http://russelmd.megastudy.net"
					   End If 

						If isDirCheck Then 
						  returnURLSite = returnURLSite & "/russel_mokdong"
						End If
				
				 Case "CD0246" '���� ����
					   If isCheckDev Then 
						 returnURLSite = "http://russelct.megastudy.net"
					   Else 
						 returnURLSite = "http://russelct.megastudy.net"
					   End If 

						If isDirCheck Then 
						  returnURLSite = returnURLSite & "/russel_ct"
						End If

				Case "CD0247" '���� ����
					   If isCheckDev Then 
						 returnURLSite = "http://russelgn.megastudy.net"
					   Else 
						 returnURLSite = "http://russelgn.megastudy.net"
					   End If 

						If isDirCheck Then 
						  returnURLSite = returnURLSite & "/russel_gn"
						End If
				
				Case "CD0248" '���� ����
					   If isCheckDev Then 
						 returnURLSite = "http://russelpc.megastudy.net"
					   Else 
						 returnURLSite = "http://russelpc.megastudy.net"
					   End If 

						If isDirCheck Then 
						  returnURLSite = returnURLSite & "/russel_pc"
						End If
				
				Case "CD0249" '���� ����
					   If isCheckDev Then 
						 returnURLSite = "http://russelyt.megastudy.net"
					   Else 
						 returnURLSite = "http://russelyt.megastudy.net"
					   End If 

						If isDirCheck Then 
						  returnURLSite = returnURLSite & "/russel_yt"
						End If

				Case "CD0250" '���� ��õ
					   If isCheckDev Then 
						 returnURLSite = "http://russelbc.megastudy.net"
					   Else 
						 returnURLSite = "http://russelbc.megastudy.net"
					   End If 

						If isDirCheck Then 
						  returnURLSite = returnURLSite & "/russel_bc"
						End If
				Case "RUSSELINTRO" '���� INTRO
					   If isCheckDev Then 
						 returnURLSite = "http://russel.megastudy.net"
					   Else 
						 returnURLSite = "http://russel.megastudy.net"
					   End If 

						If isDirCheck Then 
						  returnURLSite = returnURLSite & "/intro"
						End If


				Case "CD0251" '��õ �����
					If isCheckDev Then 
						returnURLSite = "http://tbucheon.megastudy.net"
					Else 
						returnURLSite = "http://bucheon.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/bucheon/nsu"
					End If 

				Case "CD0253" '�д� �����
					If isCheckDev Then 
						returnURLSite = "http://tbundang.megastudy.net"
					Else 
						returnURLSite = "http://bundang.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/bundang/nsu"
					End If 


				Case "CD0255" '�ϻ� �����
					If isCheckDev Then 
						returnURLSite = "http://tilsan.megastudy.net"
					Else 
						returnURLSite = "http://ilsan.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/ilsan/nsu"
					End If 


				Case "CD0259" '���ְ���
					If isCheckDev Then
						returnURLSite = "http://tyangju.megastudy.net"
					Else
						returnURLSite = "http://yangju.megastudy.net"
					End If

					If isDirCheck Then
						returnURLSite = returnURLSite & "/yangju/jaehak"
					End If

				Case "CD0261" '��������
					If isCheckDev Then
						returnURLSite = "http://tunjung.megastudy.net"
					Else
						returnURLSite = "http://unjung.megastudy.net"
					End If

					If isDirCheck Then
						returnURLSite = returnURLSite & "/unjung/jaehak"
					End If

				Case "CD0276" '���������
					If isCheckDev Then 
						returnURLSite = "http://tsongpa.megastudy.net"
					Else 
						returnURLSite = "http://songpa.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/songpa/nsu"
					End If 

				Case "CD0280" '�ȼ����
					If isCheckDev Then 
						returnURLSite = "http://tansung.megastudy.net"
					Else 
						returnURLSite = "http://ansung.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/ansung"
					End If 

				Case Else '�� �̿��ǰ�� 
					If isCheckDev Then 
						returnURLSite = "http://tcampus.megastudy.net"
					Else 
						returnURLSite = "http://campus.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/campus"
					End If 

			End Select

		End If

		GetCampusDetailUrlSiteMGC = returnURLSite 
	End Function

	' �п��� ����Ʈ�ּ� ��������  E

	' �п��� ���� index.asp ������ ���� üũ  S
	Function GetCampusIsIndexMGC()
		returnVal = False 
		' ������ ��ũ ����
		gMGC_SiteCurrentURL = Request.ServerVariables("PATH_INFO")
		gMGC_arrgSiteCurrentURL = Split(gMGC_SiteCurrentURL, "/")
		gMGC_SiteMenu_1th = LCase(gMGC_arrgSiteCurrentURL(1))
		'If UBound(gMGC_arrgSiteCurrentURL) > 1 Then
		'	gMGC_SiteMenu_2nd = LCase(gMGC_arrgSiteCurrentURL(2))
		'End If
		'If UBound(gMGC_arrgSiteCurrentURL) > 2 Then
		'	gMGC_SiteMenu_3rd = LCase(gMGC_arrgSiteCurrentURL(3))
		'End If 
		'If UBound(gMGC_arrgSiteCurrentURL) > 3 Then
		'	gMGC_SiteMenu_4th = LCase(gMGC_arrgSiteCurrentURL(4))
		'End If 

		'Index ���� üũ 
		If gMGC_SiteMenu_1th = "index.asp"  Then 
			returnVal = True  
		End If 

		GetCampusIsIndexMGC = returnVal
	End Function 
	' �п��� ���� index.asp ������ ���� üũ E


	' �׽�Ʈ � üũ   S
	Function GetCampusBaseUrlMGC()
		returnVal = "" 
		If Request.ServerVariables("LOCAL_ADDR") = "10.1.3.10" Then  '���߼��� üũ
				returnVal = "http://tcampus.megastudy.net"
		Else
				returnVal = "http://campus.megastudy.net"
		End If	

		GetCampusBaseUrlMGC = returnVal
	End Function 
	' �׽�Ʈ � üũ E



	' campus �� �ƴ� �� ������ ����Ʈ ���� üũ   S
	Function GetCampusIsDMMGC()
		gMGC_Sub_RequestServerName = LCase(Request.ServerVariables("SERVER_NAME"))  '������ �̸�
		gMGC_Sub_ServerName = "http://" & RequestServerName	
		returnVal = False 
		If gMGC_Sub_RequestServerName="campus.megastudy.net" Or gMGC_Sub_RequestServerName="campus1.megastudy.net" Or gMGC_Sub_RequestServerName="campus2.megastudy.net" Or gMGC_Sub_RequestServerName="campus3.megastudy.net" Or gMGC_Sub_RequestServerName="campus5.megastudy.net" Or gMGC_Sub_RequestServerName="amsdev.megastudy.net" Or gMGC_Sub_RequestServerName="tcampus.megastudy.net" Or gMGC_Sub_RequestServerName="mcampus.megastudy.net" Or gMGC_Sub_RequestServerName="mcampus1.megastudy.net" Or gMGC_Sub_RequestServerName="mcampus2.megastudy.net" Or gMGC_Sub_RequestServerName="mcampus3.megastudy.net" Or gMGC_Sub_RequestServerName="mcampus5.megastudy.net" Or gMGC_Sub_RequestServerName="tmcampus.megastudy.net" Then 		
			returnVal = False
		Else		
			returnVal = True 
		End If 

		GetCampusIsDMMGC = returnVal
	End Function 
	' campus �� �ƴ� �� ������ ����Ʈ ���� üũ E

	'������ �̵� S

	If VARS("clear") = "Y" And GetCampusIsIndexMGC() Then  ' index.asp �������̰� clear=Y �Ķ�����̸�  
			
			If VARS("clear") = "Y" And ( GetCampusIsDMMGC() ) Then ' Y �̰� �ٸ� ���� ķ�۽� �ּ��̸� �̵� 
				Response.redirect  GetCampusBaseUrlMGC()
				Response.End
			End If

	ElseIf ValueCheck( VARS("refer") ) And GetCampusIsIndexMGC() Then  ' index.asp �������̰� refer �Ķ���Ͱ� ������ �ش� �������� �̵�   
				
		If GetCampusIsDMMGC() Then 		 ' �ٸ� ���� ķ�۽� �ּ��̸� �̵� 
			Response.redirect  GetCampusBaseUrlMGC() & VARS("refer")
			Response.End
		End If 
		
	ElseIf GetCampusIsDMMGC() And GetCampusIsIndexMGC() Then  ' campus �� �ƴ� �� ������ ����Ʈ�̸�  index.asp �������̸� �� �������п����� �̵� 
		'�ش� �� �п�����Ʈ �̵�
		If ValueCheck(GetCampusVarsSiteMGC("_CAMPUS_DIRSITE_")) Then
			If LCase(Request.ServerVariables("SERVER_NAME")) = "localhost" Then	
				' �����϶��� �����̷�Ʈ �Ƚ�Ŵ(2020.03.10)
			Else
				Response.Redirect GetCampusVarsSiteMGC("_CAMPUS_DIRSITE_")
				Response.End
			End If
		End If 
	
	ElseIf (Not GetCampusIsDMMGC() ) And GetCampusIsIndexMGC() Then  'campus ����Ʈ �̰� index.asp ������ �̸� 

	Else 

	End If 

	gMGC_Campus_Title = GetCampusVarsSiteMGC("_CAMPUS_TITLE_")

	'������ �̵� E
	
	'������ �б� ó�� E
%>
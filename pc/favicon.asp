<%
'##################################################################
' ���ϼ��� : ���������� �������� ���Ǵ� ������
'            ������Ʈ ������ ���� ���
' ����     : 2015/11/13 - ����ȣ
' ����   : ���������� ���� ��� ? ���� ���� ������,
'            ������� ������������ ���ο� ������ �� ó�� ���̰� �Ѵ�
'##################################################################
%>
<% 
	'���̹� ��Ÿ üũ �������� �Լ� �������� �Լ� S 
	Function GetCampusNaverCheckMeta()
		
		isNaverCheck_CampusDetailCode = "0" 
		detailNaverCheck_CampusDetailCode = "9d6d7d0820f3efbf43bd3fd3ba73ca6f7449e3b5"

		NaverCheck_RequestServerName = LCase(Request.ServerVariables("SERVER_NAME"))  '������ �̸�
		
		'����Ʈ ���� ���� S
		
		Select Case NaverCheck_RequestServerName

			'Campus 
			Case "campus.megastudy.net", "campus1.megastudy.net", "campus2.megastudy.net", "campus5.megastudy.net", "amsdev.megastudy.net" ,"tcampus.megastudy.net", "mcampus.megastudy.net", "mcampus1.megastudy.net", "mcampus2.megastudy.net", "mcampus5.megastudy.net" , "tmcampus.megastudy.net", "campus3.megastudy.net", "mcampus3.megastudy.net"
	   
				NaverCheck_CampusDetailCode = ""
				detailNaverCheck_CampusDetailCode = "9d6d7d0820f3efbf43bd3fd3ba73ca6f7449e3b5"
			
			'���� �����	http://gangnam.megastudy.net	http://mgangnam.megastudy.net
			Case "gangnam.megastudy.net", "mgangnam.megastudy.net", "tgangnam.megastudy.net", "tmgangnam.megastudy.net"

				NaverCheck_CampusDetailCode = "CD0206"
				detailNaverCheck_CampusDetailCode = "40010b1608b9242b424c2eecb1c110aeb3839cc6"
			
			'���� ��.����	http://gangnams.megastudy.net	http://mgangnams.megastudy.net
			Case "gangnams.megastudy.net", "mgangnams.megastudy.net", "tgangnams.megastudy.net", "tmgangnams.megastudy.net"
			
				NaverCheck_CampusDetailCode = "CD0205"
				
			'���� �����	http://gangdong.megastudy.net	http://mgangdong.megastudy.net
			Case "gangdong.megastudy.net", "mgangdong.megastudy.net", "tgangdong.megastudy.net", "tmgangdong.megastudy.net"
						
				NaverCheck_CampusDetailCode = "CD0240"

			'���� ��.��.��	http://gangdongs.megastudy.net	http://mgangdongs.megastudy.net
			Case "gangdongs.megastudy.net", "mgangdongs.megastudy.net", "tgangdongs.megastudy.net", "tmgangdongs.megastudy.net"
				NaverCheck_CampusDetailCode = "CD0215"
			
			'���� �����	http://gangbuk.megastudy.net	http://mgangbuk.megastudy.net
			Case "gangbuk.megastudy.net", "mgangbuk.megastudy.net", "tgangbuk.megastudy.net", "tmgangbuk.megastudy.net"
			
				NaverCheck_CampusDetailCode = "CD0210"
				detailNaverCheck_CampusDetailCode = "cbf633f2c2aecda3a27e88652667839abf825c93"

			'���� ����	http://gangbuks.megastudy.net	http://mgangbuks.megastudy.net
			Case "gangbuks.megastudy.net", "mgangbuks.megastudy.net", "tgangbuks.megastudy.net", "tmgangbuks.megastudy.net"
			
				NaverCheck_CampusDetailCode = "CD0209"
				detailNaverCheck_CampusDetailCode = "82c487da1c8ef63a7b0562d1d3d16a75f5753ca1"

			'�뷮�� �����	http://noryangjin.megastudy.net	http://mnoryangjin.megastudy.net
			Case "noryangjin.megastudy.net", "mnoryangjin.megastudy.net", "tnoryangjin.megastudy.net", "tmnoryangjin.megastudy.net"
							
				NaverCheck_CampusDetailCode = "CD0211"
				detailNaverCheck_CampusDetailCode = "2d723a4927de85db9188120a8df7a2310dd9a670"

			'�뷮�� �ܰ�	http://nrj.megastudy.net	http://mnrj.megastudy.net
			Case "nrj.megastudy.net", "mnrj.megastudy.net","tnrj.megastudy.net", "tmnrj.megastudy.net"
								
				NaverCheck_CampusDetailCode = "CD0212"

			'���� �����	http://seocho.megastudy.net	http://mseocho.megastudy.net
			Case "seocho.megastudy.net", "mseocho.megastudy.net", "tseocho.megastudy.net", "tmseocho.megastudy.net"
							
				NaverCheck_CampusDetailCode = "CD0208"
				detailNaverCheck_CampusDetailCode = "dcf7f0020252811c17ed535a57eedfbe55ef7571"

			'���� ����	http://seochos.megastudy.net	http://mseochos.megastudy.net
			Case "seochos.megastudy.net", "mseochos.megastudy.net", "tseochos.megastudy.net", "tmseochos.megastudy.net"
				
				NaverCheck_CampusDetailCode = "CD0207"
				
			'���� �����	http://seongbuk.megastudy.net	http://mseongbuk.megastudy.net
			Case "seongbuk.megastudy.net", "mseongbuk.megastudy.net", "tseongbuk.megastudy.net", "tmseongbuk.megastudy.net"

				NaverCheck_CampusDetailCode = "CD0239"
				
			'���� ��.��.��	http://seongbuks.megastudy.net	 http://mseongbuks.megastudy.net
			Case "seongbuks.megastudy.net", "mseongbuks.megastudy.net", "tseongbuks.megastudy.net", "tmseongbuks.megastudy.net"
				
				NaverCheck_CampusDetailCode = "CD0214"
				detailNaverCheck_CampusDetailCode = "ef24c3e46a702277ac0787fa8afcebb03e836737"
			
			'���� ������Թ�	http://sinchon.megastudy.net	http://msinchon.megastudy.net
			Case "sinchon.megastudy.net", "msinchon.megastudy.net", "tsinchon.megastudy.net", "tmsinchon.megastudy.net"

				NaverCheck_CampusDetailCode = "CD0213"
				detailNaverCheck_CampusDetailCode = "c864f0d1a48a920d36d14e16d45e7ab54c6a555d"

			'���� �����	http://pyeongchon.megastudy.net	http://mpyeongchon.megastudy.net
			Case "pyeongchon.megastudy.net", "mpyeongchon.megastudy.net" , "tpyeongchon.megastudy.net", "tmpyeongchon.megastudy.net"

				NaverCheck_CampusDetailCode = "CD0217"
				detailNaverCheck_CampusDetailCode = "0820ea8324c4515e31b5eb346d086676a5c3af47"

			'���� ����	http://pyeongchons.megastudy.net	http://mpyeongchons.megastudy.net
			Case "pyeongchons.megastudy.net", "mpyeongchons.megastudy.net","tpyeongchons.megastudy.net", "tmpyeongchons.megastudy.net"

				NaverCheck_CampusDetailCode = "CD0216"

			'���� ���� ������Թ�	http://yangji.megastudy.net	http://myangji.megastudy.net
			Case "yangji.megastudy.net", "myangji.megastudy.net", "tyangji.megastudy.net", "tmyangji.megastudy.net"	

				NaverCheck_CampusDetailCode = "CD0220"
				detailNaverCheck_CampusDetailCode = "dbd011ceaa6cb26988c49e81a65fe47919a639ab"

			'���� �Ű�/���ʱ�� ������Թ�	http://yangjim.megastudy.net	http://myangjim.megastudy.net
			Case "yangjim.megastudy.net", "myangjim.megastudy.net"	, "tyangjim.megastudy.net", "tmyangjim.megastudy.net", "seochob.megastudy.net", "mseochob.megastudy.net"	, "tseochob.megastudy.net", "tmseochob.megastudy.net"	

				NaverCheck_CampusDetailCode = "CD0243"
				detailNaverCheck_CampusDetailCode = "f426f87682d3a84296ee912a744592765dd9425a"

			'��õ	������Թ�	http://bucheon.megastudy.net
			Case "bucheon.megastudy.net", "mbucheon.megastudy.net", "tbucheon.megastudy.net", "tmbucheon.megastudy.net"	

				NaverCheck_CampusDetailCode = "CD0251"
				detailNaverCheck_CampusDetailCode = "0ade76b50f0293213b75e09fadf8884ebfae23c5"
			
			'�ϻ�	������Թ�	http://ilsan.megastudy.net
			Case "ilsan.megastudy.net", "milsan.megastudy.net", "tilsan.megastudy.net", "tmilsan.megastudy.net"	

				NaverCheck_CampusDetailCode = "CD0255"
				detailNaverCheck_CampusDetailCode = "7b5099ed5e72ad2a983d07fe13196087d7f7bd4b"

			'�д�	������Թ�	http://bundang.megastudy.net
			Case "bundang.megastudy.net", "mbundang.megastudy.net", "tbundang.megastudy.net", "tmbundang.megastudy.net"	

				NaverCheck_CampusDetailCode = "CD0253"
				detailNaverCheck_CampusDetailCode = "be67589cde709fa3569b934ba920a74e6d66edaa"

			'������ġ http://campus.magastudy.net/russel/
            Case "russeldc.megastudy.net", "mrusseldc.megastudy.net"

                 NaverCheck_CampusDetailCode = "CD0241"

           '�����д� http://campus.magastudy.net/russel_bundang/
           Case "russelbd.megastudy.net", "mrusselbd.megastudy.net"

                NaverCheck_CampusDetailCode = "CD0244"

           '������ http://campus.magastudy.net/russel_mokdong/
            Case "russelmd.megastudy.net", "mrusselmd.megastudy.net"

                NaverCheck_CampusDetailCode = "CD0245"

			'���� ��.��.��	http://yangju.megastudy.net	 http://myangju.megastudy.net
			Case "yangju.megastudy.net", "myangju.megastudy.net", "tyangju.megastudy.net", "tmyangju.megastudy.net"
				
				NaverCheck_CampusDetailCode = "CD0259"

			'���� �����	http://songpa.megastudy.net	http://msongpa.megastudy.net
			Case "songpa.megastudy.net", "msongpa.megastudy.net", "tsongpa.megastudy.net", "tmsongpa.megastudy.net"
						
				NaverCheck_CampusDetailCode = "CD0276"
				detailNaverCheck_CampusDetailCode = "005c26b930b430e9b0af9598a47395b4df02abe2"

			'�ȼ����	http://ansung.megastudy.net	http://mansung.megastudy.net
			Case "ansung.megastudy.net", "mansung.megastudy.net", "tansung.megastudy.net", "tmansung.megastudy.net"	
						
				NaverCheck_CampusDetailCode = "CD0280"				

			Case Else

				NaverCheck_CampusDetailCode = "" 
			
		End Select

		'����Ʈ ���� ���� E

		'���� �ް����͵��п� ����� CD0206	'���� �ް����͵��п� �߰��� CD0205	'���ϸް����͵��п� ����� CD0210	'���ϸް����͵��п� ���� CD0209	'�뷮�� �ް����͵��п� ����� CD0211	'���� �ް����͵��п� CD0213	'���� �ް����͵��п� CD0217	'���� �ް����͵��п� ���� CD0216	'���� �ް����͵��п� ����� CD0240	'���� �ް����͵��п� ���߰�� CD0215 '���� �ް����͵��п� ���߰�� CD0276 
		'<meta name="naver-site-verification" content="6f8d9536b461cf0fb667a1a89afac4a92b8ae961"/> ���
		'�׿�  <meta name="naver-site-verification" content="8e19e1ad08a3c6b9abee62054f8eedddd706f119"/>
		' ���Ĵ� 005c26b930b430e9b0af9598a47395b4df02abe2 �ڵ�� ���� (2020-12-10)

		If NaverCheck_CampusDetailCode = "CD0206" Or NaverCheck_CampusDetailCode = "CD0205" Or NaverCheck_CampusDetailCode = "CD0210" Or NaverCheck_CampusDetailCode = "CD0209" Or NaverCheck_CampusDetailCode = "CD0211" Or NaverCheck_CampusDetailCode = "CD0213" Or NaverCheck_CampusDetailCode = "CD0217" Or NaverCheck_CampusDetailCode = "CD0216" Or NaverCheck_CampusDetailCode = "CD0240" Or NaverCheck_CampusDetailCode = "CD0215" Or NaverCheck_CampusDetailCode = "CD0259" Or NaverCheck_CampusDetailCode = "CD0251" Or NaverCheck_CampusDetailCode = "CD0253" Or NaverCheck_CampusDetailCode = "CD0208" Or NaverCheck_CampusDetailCode = "CD0207"  Or NaverCheck_CampusDetailCode = "CD0239"  Or NaverCheck_CampusDetailCode = "CD0214" Or NaverCheck_CampusDetailCode = "CD0255" Or NaverCheck_CampusDetailCode = "CD0220" Or NaverCheck_CampusDetailCode = "CD0243" Then 
			'isNaverCheck_CampusDetailCode = "1"
		ElseIf NaverCheck_CampusDetailCode = "CD0276" Then
			'isNaverCheck_CampusDetailCode = "2"
		Else 
			'isNaverCheck_CampusDetailCode = "0" 
		End If 

		'GetCampusNaverCheckMeta =  isNaverCheck_CampusDetailCode
		' �п����� ��Ÿ�±� �ڵ尪 ����� (2021-03-16)
		GetCampusNaverCheckMeta =  detailNaverCheck_CampusDetailCode
	End Function 
	'���̹� ��Ÿ üũ �������� �Լ� E




	'���̹� ��Ÿ ���� �������� �Լ� �������� �Լ� S 
	Function GetCampusNaverDescMeta()
		
		NaverCheck_CampusDetailDesc = ""

		NaverCheck_RequestServerName = LCase(Request.ServerVariables("SERVER_NAME"))  '������ �̸�
		
		'����Ʈ ���� ���� S
		
		Select Case NaverCheck_RequestServerName

			'Campus 
			Case "campus.megastudy.net", "campus1.megastudy.net", "campus2.megastudy.net", "campus5.megastudy.net", "amsdev.megastudy.net" ,"tcampus.megastudy.net", "mcampus.megastudy.net", "mcampus1.megastudy.net", "mcampus2.megastudy.net", "mcampus5.megastudy.net" , "tmcampus.megastudy.net"
	   
				NaverCheck_CampusDetailDesc = "���Լ����� ��� ����, ���� ����, ����/�������п�, �Խ�/��������, ���غ� �����հݽý���, ������������"
			
			'���� �����	http://gangnam.megastudy.net	http://mgangnam.megastudy.net
			Case "gangnam.megastudy.net", "mgangnam.megastudy.net", "tgangnam.megastudy.net", "tmgangnam.megastudy.net"

				NaverCheck_CampusDetailDesc = "����/��ġ/���� ����п�, ����/�Խ�����, ������������, ������"
			
			'���� ��.����	http://gangnams.megastudy.net	http://mgangnams.megastudy.net
			Case "gangnams.megastudy.net", "mgangnams.megastudy.net", "tgangnams.megastudy.net", "tmgangnams.megastudy.net"
			
				NaverCheck_CampusDetailDesc = "������ ��ġ�� ��ġ, ���л������п�, ����/���� ���, ��/���� ���չ�, �Խ��п�"
				
			'���� �����	http://gangdong.megastudy.net	http://mgangdong.megastudy.net
			Case "gangdong.megastudy.net", "mgangdong.megastudy.net", "tgangdong.megastudy.net", "tmgangdong.megastudy.net"
						
				NaverCheck_CampusDetailDesc = "����/����/��� ����п�, ����/�Խ�����, ������������, ���丵 ���α׷�, ������Թ�"

			'���� ��.��.��	http://gangdongs.megastudy.net	http://mgangdongs.megastudy.net
			Case "gangdongs.megastudy.net", "mgangdongs.megastudy.net", "tgangdongs.megastudy.net", "tmgangdongs.megastudy.net"

				NaverCheck_CampusDetailDesc = "���߰� ���л� ���� �п�, �������, �����ü��, �ܰ��� �, ����/���� �Ϻ� ���"
			
			'���� �����	http://gangbuk.megastudy.net	http://mgangbuk.megastudy.net
			Case "gangbuk.megastudy.net", "mgangbuk.megastudy.net", "tgangbuk.megastudy.net", "tmgangbuk.megastudy.net"
			
				NaverCheck_CampusDetailDesc = "���� �Խ��� �߽�! ����/���/��� ����п�, ����/�Խ�����, ������������, ������"

			'���� ����	http://gangbuks.megastudy.net	http://mgangbuks.megastudy.net
			Case "gangbuks.megastudy.net", "mgangbuks.megastudy.net", "tgangbuks.megastudy.net", "tmgangbuks.megastudy.net"
			
				NaverCheck_CampusDetailDesc = "��� ����, �����ü��, �Խ�����, ����/���� ���ô��, ����� ��赿 ��ġ"

			'�뷮�� �����	http://noryangjin.megastudy.net	http://mnoryangjin.megastudy.net
			Case "noryangjin.megastudy.net", "mnoryangjin.megastudy.net", "tnoryangjin.megastudy.net", "tmnoryangjin.megastudy.net"
							
				NaverCheck_CampusDetailDesc = "�ް����͵��п� ��� ������, ����и� �ڽ���, ������ ����, �Խ����� ����"

			'�뷮�� �ܰ�	http://nrj.megastudy.net	http://mnrj.megastudy.net
			Case "nrj.megastudy.net", "mnrj.megastudy.net","tnrj.megastudy.net", "tmnrj.megastudy.net"
								
				NaverCheck_CampusDetailDesc = "���۱� �뷮���� ��ġ, �뷮�� �ܰ�"

			'���� �����	http://seocho.megastudy.net	http://mseocho.megastudy.net
			Case "seocho.megastudy.net", "mseocho.megastudy.net", "tseocho.megastudy.net", "tmseocho.megastudy.net"
							
				NaverCheck_CampusDetailDesc = "��ǥ�� �ϼ��ϴ�, �Ϻ��� N���� ����! ��������, ������������, ��������"

			'���� ����	http://seochos.megastudy.net	http://mseochos.megastudy.net
			Case "seochos.megastudy.net", "mseochos.megastudy.net", "tseochos.megastudy.net", "tmseochos.megastudy.net"
				
				NaverCheck_CampusDetailDesc = "���� ������ ������ ����! �Ǵ��Խ�����, �ڿ��� �ֻ����� ����, ������������, ������Թ�"
				
			'���� �����	http://seongbuk.megastudy.net	http://mseongbuk.megastudy.net
			Case "seongbuk.megastudy.net", "mseongbuk.megastudy.net", "tseongbuk.megastudy.net", "tmseongbuk.megastudy.net"

				NaverCheck_CampusDetailDesc = "����/����/���빮 ����п�, ����/�Խ�����, ������, ������������, �������� �������, ������Թ�"
				
			'���� ��.��.��	http://seongbuks.megastudy.net	 http://mseongbuks.megastudy.net
			Case "seongbuks.megastudy.net", "mseongbuks.megastudy.net", "tseongbuks.megastudy.net", "tmseongbuks.megastudy.net"
				
				NaverCheck_CampusDetailDesc = "2023 ����� �Ǵ� �����հ�. ��� �ܰ� ������Ž. ��� �б��� �������Ŵܰ�. ��3 ����&���� ALL KILL ����. �ߵ�/�ʵ� ����/����/����"
			
			'���� ������Թ�	http://sinchon.megastudy.net	http://msinchon.megastudy.net
			Case "sinchon.megastudy.net", "msinchon.megastudy.net", "tsinchon.megastudy.net", "tmsinchon.megastudy.net"

				NaverCheck_CampusDetailDesc = "������ �Խ� ���� ���� ������! ȫ��/����/���� ����п�, ����/�Խ�����, ������������, ������"

			'���� �����	http://pyeongchon.megastudy.net	http://mpyeongchon.megastudy.net
			Case "pyeongchon.megastudy.net", "mpyeongchon.megastudy.net" , "tpyeongchon.megastudy.net", "tmpyeongchon.megastudy.net"

				NaverCheck_CampusDetailDesc = "������ ������, �հ��� ���� �ٸ� ����! �Ⱦ�/����/�Ȼ�/��õ ����п�, ����/�Խ�����, ������������, ������"

			'���� ����	http://pyeongchons.megastudy.net	http://mpyeongchons.megastudy.net
			Case "pyeongchons.megastudy.net", "mpyeongchons.megastudy.net","tpyeongchons.megastudy.net", "tmpyeongchons.megastudy.net"

				NaverCheck_CampusDetailDesc = "�Ⱦ�, ����  ���л������п�, ����/���� ���, ���� ���չ�, �ܰ��� �"

			'���� ��� ���� ������Թ�	http://yangji.megastudy.net	http://myangji.megastudy.net
			Case "yangji.megastudy.net", "myangji.megastudy.net", "tyangji.megastudy.net", "tmyangji.megastudy.net"	

				NaverCheck_CampusDetailDesc = "�������� ����й� ���̰� ������������, �ް����͵��� ����, ����/�Խ�����, ������������, ������, ���ν� ó�α� ����� ��ġ"

			' ���� ��� ���� �Ű� ������Թ�	http://yangjim.megastudy.net	http://myangjim.megastudy.net
			Case "yangjim.megastudy.net", "myangjim.megastudy.net"	, "tyangjim.megastudy.net", "tmyangjim.megastudy.net" ,"seochob.megastudy.net", "mseochob.megastudy.net"	, "tseochob.megastudy.net", "tmseochob.megastudy.net"		

				NaverCheck_CampusDetailDesc = "(��) ���� �ް����͵� ����п�, �ް����͵� ���� ����п�, �޵��� �հ��� ���� �ּ��� ����, ������������, ���� �ڿ��� �հ��� ����, �ڿ��迭 �л����� ������� Focus�� ���� ����, ���� Ưȭ����/����, Ž������ ���غ� ����, ������ ������� ���Ͽ�, Miracle ���� �Ǵ��" 
			
			'��õ	������Թ�	http://bucheon.megastudy.net
			Case "bucheon.megastudy.net", "mbucheon.megastudy.net", "tbucheon.megastudy.net", "tmbucheon.megastudy.net"	

				NaverCheck_CampusDetailDesc = "���Ѱ��� Ȯ���� ���! ��õ/����/����/���� ����п�, ����/�Խ�����, ������������, ������"
			
			'�ϻ�	������Թ�	http://ilsan.megastudy.net
			Case "ilsan.megastudy.net", "milsan.megastudy.net", "tilsan.megastudy.net", "tmilsan.megastudy.net"	

				NaverCheck_CampusDetailDesc = "������������������ �߽�! �ϻ�/����/���� ����п�, ����/�Խ�����, ������������, ������"

			'�д�	������Թ�	http://bundang.megastudy.net
			Case "bundang.megastudy.net", "mbundang.megastudy.net", "tbundang.megastudy.net", "tmbundang.megastudy.net"	

				NaverCheck_CampusDetailDesc = "���غ� ���� �н�, ö���� �л� ����! ����/����/����/���� ����п�, ����/�Խ�����, ������������, ������"

			'������ġ http://campus.magastudy.net/russel/
            Case "russeldc.megastudy.net", "mrusseldc.megastudy.net"

                 NaverCheck_CampusDetailDesc = ""

           '�����д� http://campus.magastudy.net/russel_bundang/
           Case "russelbd.megastudy.net", "mrusselbd.megastudy.net"

                NaverCheck_CampusDetailDesc = ""

           '������ http://campus.magastudy.net/russel_mokdong/
            Case "russelmd.megastudy.net", "mrusselmd.megastudy.net"

                NaverCheck_CampusDetailDesc = ""

			'���� ��.��.��	http://yangju.megastudy.net	 http://myangju.megastudy.net
			Case "yangju.megastudy.net", "myangju.megastudy.net", "tyangju.megastudy.net", "tmyangju.megastudy.net"
				
				NaverCheck_CampusDetailDesc = "����/������/����õ/������� ���߰� �Խ�����, ��������, �ڵ�Ư��, ����/���� �Ϻ����, ���������"
			
			'���� ��.��.��	http://unjung.megastudy.net	http://munjung.megastudy.net/
			Case "unjung.megastudy.net", "munjung.megastudy.net", "tunjung.megastudy.net", "tmunjung.megastudy.net"
				NaverCheck_CampusDetailDesc = "����/����/�������� ���߰� �Խ�����, �������, �ڵ�Ư��, ����/���� �Ϻ� ���"

			'���� �����	http://songpa.megastudy.net	http://msongpa.megastudy.net
			Case "songpa.megastudy.net", "msongpa.megastudy.net", "tsongpa.megastudy.net", "tmsongpa.megastudy.net"						
				NaverCheck_CampusDetailDesc = "���� ������ ���� �ֻ��� ���! ����/����/����/����/����/�ϳ�/������ ����п�, ����/�Խ�����, ������������, ������"

			'�ȼ���� 	http://ansung.megastudy.net	http://mansung.megastudy.net
			Case "ansung.megastudy.net", "mansung.megastudy.net", "tansung.megastudy.net", "tmansung.megastudy.net"						
				NaverCheck_CampusDetailDesc = "����� ������ ������ ������! �ȼ����� �������, ����/�������п�, �������� ����/����п�, ���ɼ��� ��� ����, ������������"

			Case Else

				NaverCheck_CampusDetailDesc = "���Լ����� ��� ����, ���� ����, ����/�������п�, �Խ�/��������, ���غ� �����հݽý���, ������������"
			
		End Select

		'����Ʈ ���� ���� E

		'���ް� ���ǰ�� ȫ�������� ó�� / �ֹμ� / 2025-01-10 17:23:50
		If LCase(request.ServerVariables("PATH_INFO")) = "/campus_common/2025/2025_omega/index.asp" Then
			NaverCheck_CampusDetailDesc = "�Ƿ°� ������ ������ ������ ���ް�! ���� ���� ���� ���� ���� �ϼ�!"
		'�޴��� ���ǰ�� ȫ�������� ó�� / �ֹμ� / 2025-01-15
		ElseIf LCase(request.ServerVariables("PATH_INFO")) = "/campus_common/2025/2025_premium/index.asp" Then
			NaverCheck_CampusDetailDesc = "�� ���� �� ���� ���ǰ��� �Ϻ��� ���� ���� ����, �� ��ü�� �����̾��� ���ǰ��"
		End If

		GetCampusNaverDescMeta =  NaverCheck_CampusDetailDesc
	End Function 
	'���̹� ��Ÿ ���� �������� �Լ� E
%>
<!-- favicon -->	
<link rel="shortcut icon" href="https://img.megastudy.net/campus/library/v2015_mob/asset/img/favicon.ico">
<link rel="apple-touch-icon" href="https://img.megastudy.net/campus/library/v2015_mob/asset/img/iphone.png">
<link rel="apple-touch-icon" sizes="76x76" href="https://img.megastudy.net/campus/library/v2015_mob/asset/img/apple-touch-icon-ipad.png">
<link rel="apple-touch-icon" sizes="120x120" href="https://img.megastudy.net/campus/library/v2015_mob/asset/img/apple-touch-icon-iphone4.png">

<%If LCase(request.ServerVariables("PATH_INFO")) = "/campus_common/2025/2025_omega/index.asp" Then%>
	<!--���ް� ���ǰ�� ȫ�������� ó�� / �ֹμ� / 2025-01-10 17:24:56 -->
	<meta name="author" content="�ް����͵��п�" />
	<meta name="keywords" content="�ް����͵��п�, �����п�, ����п�, �Խ�, �������п�, �ް����͵𷯼�, �ް����͵�, N��, ����, ����, ���ǰ��, ���� ���ǰ��" />
	<meta property="og:site_name" content="�ް����͵��п�">
	<meta property="og:image" content="https://img.megastudy.net/campus/library/v2015/library/campus_common/2025/2025_omega/thumb.jpg">
	<meta property="og:url" content="https://campus.megastudy.net/campus_common/2025/2025_omega/index.asp">
<%ElseIf LCase(request.ServerVariables("PATH_INFO")) = "/campus_common/2025/2025_premium/index.asp" Then%>
	<!--�޴��� ���ǰ�� ȫ�������� ó�� / �ֹμ� / 2025-01-15 -->
	<meta name="author" content="�ް����͵��п�" />
	<meta name="keywords" content="�ް����͵��п�, �����п�, ����п�, �Խ�, �������п�, �ް����͵𷯼�, �ް����͵�, N��, ����, ����, ���ǰ��, �� �����̾� ���ǰ��" />
	<meta property="og:site_name" content="�ް����͵��п�">
	<meta property="og:image" content="https://img.megastudy.net/campus/library/v2015/library/campus_common/2025/2025_premium/thumb.jpg">
	<meta property="og:url" content="https://campus.megastudy.net/campus_common/2025/2025_premium/index.asp">
<%Else%>
	<meta property="og:image" content="https://img.megastudy.net/campus/library/v2015/library/intro/main_logo.gif">
	<meta property="og:url" content="https://<%=Request.ServerVariables("SERVER_NAME")%>">
<%End If%>

<meta name="naver-site-verification" content="<%=GetCampusNaverCheckMeta()%>"/>
<meta name="description" content="<%=GetCampusNaverDescMeta()%>"/>
<meta property="og:type" content="website">
<meta property="og:title" content="<%=gMGC_Campus_Title & IIF(GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_")="CD0214", " ���߰��", "") %>">
<meta property="og:description" content="<%=GetCampusNaverDescMeta()%>">
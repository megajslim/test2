<%
	'------------------------------------------------------------------------------------------------------------
	' 점검
	'------------------------------------------------------------------------------------------------------------
	' 점검시간
'	If CDate("2018-04-12 00:00:00") <= Now() And Now() < CDate("2018-04-12 07:00:00") Then
'		Response.Redirect "/index_180412.asp"
'		Response.End
'	End If

	'도메인 분기 처리 S

	'학원별 변수 가져오기 함수 S 
	Function GetCampusVarsSiteMGC(VarsName)
		
		'변수 도메인 
		gMGC_RequestServerName = LCase(Request.ServerVariables("SERVER_NAME"))  '도메인 이름
		gMGC_ServerName = "http://" & RequestServerName	

		gMGC_RequestPathInfo = LCase(Request.ServerVariables("PATH_INFO"))	' 페이지경로
		
		'사이트 변수 설정 S
		
		'강남	재수관	http://gangnam.megastudy.net	CD0206	CD0001	http://mgangnam.megastudy.net
		'강남 중.고등관	http://gangnams.megastudy.net	CD0205	CD0001	http://mgangnams.megastudy.net
		'강동	재수관	http://gangdong.megastudy.net	CD0240	CD0198	http://mgangdong.megastudy.net
		'강동 초.중.고	http://gangdongs.megastudy.net	CD0215	CD0198	http://mgangdongs.megastudy.net
		'강북	재수관	http://gangbuk.megastudy.net	CD0210	CD0005	http://mgangbuk.megastudy.net
		'강북 고등관	http://gangbuks.megastudy.net	CD0209	CD0005	http://mgangbuks.megastudy.net
		'노량진	재수관	http://noryangjin.megastudy.net	CD0211	CD0006	http://mnoryangjin.megastudy.net
		'노량진 단과	http://nrj.megastudy.net	CD0212	CD0006	http://mnrj.megastudy.net
		'서초	재수관	http://seocho.megastudy.net	CD0208	CD0004	http://mseocho.megastudy.net
		'서초 고등관	http://seochos.megastudy.net	CD0207	CD0004	http://mseochos.megastudy.net
		'성북	재수관	http://seongbuk.megastudy.net	CD0239	CD0098	http://mseongbuk.megastudy.net
		'성북 초.중.고	http://seongbuks.megastudy.net	CD0214	CD0098	http://mseongbuks.megastudy.net
		'신촌 재수정규반	http://sinchon.megastudy.net	CD0213	CD0036	http://msinchon.megastudy.net
		'평촌	재수관	http://pyeongchon.megastudy.net	CD0217	CD0179	http://mpyeongchon.megastudy.net
		'평촌고등관	http://pyeongchons.megastudy.net	CD0216	CD0179	http://mpyeongchons.megastudy.net
		'양지 본원	재수정규반	http://yangji.megastudy.net	CD0220	CD0178	http://myangji.megastudy.net
		'양지 신관	재수정규반	http://yangjim.megastudy.net	CD0243	CD0039	http://myangjim.megastudy.net
        '부천재수정규반	http://bucheon.megastudy.net	CD0243	CD0039	http://mbucheon.megastudy.net
		'서초 메가스터디 기숙학원 	재수정규반	http://seochob.megastudy.net	CD0243	CD0039	http://mseochob.megastudy.net
		'양주 초.중.고	http://yangju.megastudy.net	CD0259	CD0258	http://myangju.megastudy.net
		'운정 초.중.고	http://unjung.megastudy.net	CD0261	CD0260	http://munjung.megastudy.net
		If IsDevSvr() = True Then
			cDev = "DEV-"
		Else
			cDev = ""
		End If

		If gMGC_RequestServerName = "localhost" Then	' 로컬에서 작업을 위해 추가(2020.03.10)

			If InStr(gMGC_RequestPathInfo, "/gangnam/nsu") > 0 Then

				'강남 재수관
				gMGC_Campus_N = "gangnam"
				gMGC_Campus_Code = "CD0001"
				gMGC_Campus_KN = "강남 메가스터디학원"
				gMGC_Campus_EN = "gangnam"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "재수관"
				gMGC_Campus_Detail_Code = "CD0206"
				gMGC_Campus_DirSite ="/gangnam/nsu"
				gMGC_Campus_Title = cDev&"강남 메가스터디학원"
				gMGC_Campus_Theme = "red"

			ElseIf InStr(gMGC_RequestPathInfo, "/gangnam/jaehak") > 0 Then

				'강남 중.고등관
				gMGC_Campus_N = "gangnam"
				gMGC_Campus_Code = "CD0001"
				gMGC_Campus_KN = "강남 메가스터디학원"
				gMGC_Campus_EN = "gangnam"
				gMGC_Campus_Kind = "jaehak"
				gMGC_Campus_Kind_KN = "고등관"
				gMGC_Campus_Detail_Code = "CD0205"
				gMGC_Campus_DirSite ="/gangnam/jaehak"
				gMGC_Campus_Title = cDev&"강남 메가스터디학원 중고등관"
				gMGC_Campus_Theme = "blue"

			ElseIf InStr(gMGC_RequestPathInfo, "/gangdong/nsu") > 0 Then
				
				'강동 재수관
				gMGC_Campus_N = "gangdong"
				gMGC_Campus_Code = "CD0198"
				gMGC_Campus_KN = "강동 메가스터디학원"
				gMGC_Campus_EN = "gangdong"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "재수관"
				gMGC_Campus_Detail_Code = "CD0240"
				gMGC_Campus_DirSite ="/gangdong/nsu"
				gMGC_Campus_Title = cDev&"강동 메가스터디학원"
				gMGC_Campus_Theme = "red"

			ElseIf InStr(gMGC_RequestPathInfo, "/gangdong/jaehak") > 0 Then

				'강동 초.중.고
				gMGC_Campus_N = "gangdong"
				gMGC_Campus_Code = "CD0198"
				gMGC_Campus_KN = "강동 메가스터디학원"
				gMGC_Campus_EN = "gangdong"
				gMGC_Campus_Kind = "jaehak"
				gMGC_Campus_Kind_KN = "초중고관"
				gMGC_Campus_Detail_Code = "CD0215"
				gMGC_Campus_DirSite ="/gangdong/jaehak"
				gMGC_Campus_Title = cDev&"강동 메가스터디학원 초중고관"
				gMGC_Campus_Theme = "blue" 

			ElseIf InStr(gMGC_RequestPathInfo, "/gangbuk/nsu") > 0 Then

				'강북 재수관
				gMGC_Campus_N = "gangbuk"
				gMGC_Campus_Code = "CD0005"
				gMGC_Campus_KN = "강북 메가스터디학원"
				gMGC_Campus_EN = "gangbuk"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "재수관"
				gMGC_Campus_Detail_Code = "CD0210"
				gMGC_Campus_DirSite ="/gangbuk/nsu"
				gMGC_Campus_Title = cDev&"강북 메가스터디학원"
				gMGC_Campus_Theme = "red"

			ElseIf InStr(gMGC_RequestPathInfo, "/gangbuk/jaehak") > 0 Then

				'강북 고등관
				gMGC_Campus_N = "gangbuk"
				gMGC_Campus_Code = "CD0005"
				gMGC_Campus_KN = "강북 메가스터디학원"
				gMGC_Campus_EN = "gangbuk"
				gMGC_Campus_Kind = "jaehak"
				gMGC_Campus_Kind_KN = "고등 종합·단과학원"
				gMGC_Campus_Detail_Code = "CD0209"
				gMGC_Campus_DirSite ="/gangbuk/jaehak"
				gMGC_Campus_Title = cDev&"강북 메가스터디학원 고등관"
				gMGC_Campus_Theme = "blue" 

			ElseIf InStr(gMGC_RequestPathInfo, "/noryangjin/nsu") > 0 Then

				'노량진 재수관
				gMGC_Campus_N = "noryangjin"
				gMGC_Campus_Code = "CD0006"
				gMGC_Campus_KN = "노량진 메가스터디학원"
				gMGC_Campus_EN = "noryangjin"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "재수관"
				gMGC_Campus_Detail_Code = "CD0211"
				gMGC_Campus_DirSite ="/noryangjin/nsu"
				gMGC_Campus_Title = cDev&"노량진 메가스터디학원"
				gMGC_Campus_Theme = "red"

			ElseIf InStr(gMGC_RequestPathInfo, "/noryangjin/danka") > 0 Then

				'노량진 단과
				gMGC_Campus_N = "noryangjin"
				gMGC_Campus_Code = "CD0006"
				gMGC_Campus_KN = "노량진 메가스터디학원"
				gMGC_Campus_EN = "noryangjin"
				gMGC_Campus_Kind = "danka"
				gMGC_Campus_Kind_KN = "단과"
				gMGC_Campus_Detail_Code = "CD0212"
				gMGC_Campus_DirSite ="/noryangjin/danka"
				gMGC_Campus_Title = cDev&"노량진 메가스터디학원 단과"
				gMGC_Campus_Theme = "blue" 

			ElseIf InStr(gMGC_RequestPathInfo, "/seocho/nsu") > 0 Then

				'서초 재수관
				gMGC_Campus_N = "seocho"
				gMGC_Campus_Code = "CD0004"
				gMGC_Campus_KN = "서초 메가스터디학원"
				gMGC_Campus_EN = "seocho"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "재수관"
				gMGC_Campus_Detail_Code = "CD0208"
				gMGC_Campus_DirSite ="/seocho/nsu"
				gMGC_Campus_Title = cDev&"서초 메가스터디학원"
				gMGC_Campus_Theme = "red"

			ElseIf InStr(gMGC_RequestPathInfo, "/seocho/jaehak") > 0 Then

				'서초 고등관
				gMGC_Campus_N = "seocho"
				gMGC_Campus_Code = "CD0004"
				gMGC_Campus_KN = "서초 메가스터디학원"
				gMGC_Campus_EN = "seocho"
				gMGC_Campus_Kind = "jaehak"
				gMGC_Campus_Kind_KN = "고등관"
				gMGC_Campus_Detail_Code = "CD0207"
				gMGC_Campus_DirSite ="/seocho/jaehak"
				gMGC_Campus_Title = cDev&"서초 메가스터디학원 고등관"
				gMGC_Campus_Theme = "blue"

			ElseIf InStr(gMGC_RequestPathInfo, "/seongbuk/nsu") > 0 Then
				
				'성북 재수관
				gMGC_Campus_N = "seongbuk"
				gMGC_Campus_Code = "CD0098"
				gMGC_Campus_KN = "성북 메가스터디학원"
				gMGC_Campus_EN = "seongbuk"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "재수관"
				gMGC_Campus_Detail_Code = "CD0239"
				gMGC_Campus_DirSite ="/seongbuk/nsu"
				gMGC_Campus_Title = cDev&"성북 메가스터디학원"
				gMGC_Campus_Theme = "red" 

			ElseIf InStr(gMGC_RequestPathInfo, "/seongbuk/jaehak") > 0 Then

				'성북 초.중.고
				gMGC_Campus_N = "seongbuk"
				gMGC_Campus_Code = "CD0098"
				gMGC_Campus_KN = "성북 메가스터디학원"
				gMGC_Campus_EN = "seongbuk"
				gMGC_Campus_Kind = "jaehak"
				gMGC_Campus_Kind_KN = "초중고관"
				gMGC_Campus_Detail_Code = "CD0214"
				gMGC_Campus_DirSite ="/seongbuk/jaehak"
				gMGC_Campus_Title = cDev&"성북 메가스터디학원 초중고관"
				gMGC_Campus_Theme = "blue" 

			ElseIf InStr(gMGC_RequestPathInfo, "/sinchon") > 0 Then

				'신촌 재수정규반
				gMGC_Campus_N = "sinchon"
				gMGC_Campus_Code = "CD0036"
				gMGC_Campus_KN = "신촌 메가스터디학원"
				gMGC_Campus_EN = "sinchon"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "재수/단과"
				gMGC_Campus_Detail_Code = "CD0213"
				gMGC_Campus_DirSite ="/sinchon"
				gMGC_Campus_Title = cDev&"신촌 메가스터디학원"
				gMGC_Campus_Theme = "red"

			ElseIf InStr(gMGC_RequestPathInfo, "/pyeongchon/nsu") > 0 Then
				
				'평촌 재수관
				gMGC_Campus_N = "pyeongchon"
				gMGC_Campus_Code = "CD0179"
				gMGC_Campus_KN = "평촌 메가스터디학원"
				gMGC_Campus_EN = "pyeongchon"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "재수관"
				gMGC_Campus_Detail_Code = "CD0217"
				gMGC_Campus_DirSite ="/pyeongchon/nsu"
				gMGC_Campus_Title = cDev&"평촌 메가스터디학원"
				gMGC_Campus_Theme = "red" 

			ElseIf InStr(gMGC_RequestPathInfo, "/pyeongchon/jaehak") > 0 Then

				'평촌 고등관
				gMGC_Campus_N = "pyeongchon"
				gMGC_Campus_Code = "CD0179"
				gMGC_Campus_KN = "평촌 메가스터디학원"
				gMGC_Campus_EN = "pyeongchon"
				gMGC_Campus_Kind = "jaehak"
				gMGC_Campus_Kind_KN = "고등관"
				gMGC_Campus_Detail_Code = "CD0216"
				gMGC_Campus_DirSite ="/pyeongchon/jaehak"
				gMGC_Campus_Title = cDev&"평촌 메가스터디학원 고등관"
				gMGC_Campus_Theme = "blue" 

			ElseIf InStr(gMGC_RequestPathInfo, "/yangji") > 0 Then

				'양지 본원 재수정규반
				gMGC_Campus_N = "yangji"
				gMGC_Campus_Code = "CD0178"
				gMGC_Campus_KN = "최상위권 전문관 러셀 기숙학원"
				gMGC_Campus_EN = "yangji"
				gMGC_Campus_Kind = "gisuk"
				gMGC_Campus_Kind_KN = ""
				gMGC_Campus_Detail_Code = "CD0220"
				gMGC_Campus_DirSite ="/yangji"
				gMGC_Campus_Title = cDev&"최상위권 전문관 러셀 기숙학원"
				gMGC_Campus_Theme = "beige"

			ElseIf InStr(gMGC_RequestPathInfo, "/yangjim") > 0 Then
				
				'서초기숙 재수정규반
				gMGC_Campus_N = "yangjim"
				gMGC_Campus_Code = "CD0039"
				gMGC_Campus_KN = "자연계 전문관 러셀 기숙학원"
				gMGC_Campus_EN = "yangjim"
				gMGC_Campus_Kind = "gisuk"
				gMGC_Campus_Kind_KN = ""
				gMGC_Campus_Detail_Code = "CD0243"
				gMGC_Campus_DirSite ="/yangjim"
				gMGC_Campus_Title = cDev&"자연계 전문관 러셀 기숙학원"
				gMGC_Campus_Theme = "beige"

			ElseIf InStr(gMGC_RequestPathInfo, "/bucheon/nsu") > 0 Then

				'부천 재수정규반
				gMGC_Campus_N = "bucheon"
				gMGC_Campus_Code = "CD0250"
				gMGC_Campus_KN = "부천 메가스터디학원"
				gMGC_Campus_EN = "bucheon"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "재수관"
				gMGC_Campus_Detail_Code = "CD0251"
				gMGC_Campus_DirSite ="/bucheon/nsu"
				gMGC_Campus_Title = cDev&"부천 메가스터디학원"
				gMGC_Campus_Theme = "red"

			ElseIf InStr(gMGC_RequestPathInfo, "/bundang/nsu") > 0 Then

				'분당 재수정규반
				gMGC_Campus_N = "bundang"
				gMGC_Campus_Code = "CD0252"
				gMGC_Campus_KN = "분당 메가스터디학원"
				gMGC_Campus_EN = "bundang"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "재수관"
				gMGC_Campus_Detail_Code = "CD0253"
				gMGC_Campus_DirSite ="/bundang/nsu"
				gMGC_Campus_Title = cDev&"분당 메가스터디학원"
				gMGC_Campus_Theme = "red" 

			ElseIf InStr(gMGC_RequestPathInfo, "/ilsan/nsu") > 0 Then

				'일산 재수정규반
				gMGC_Campus_N = "ilsan"
				gMGC_Campus_Code = "CD0254"
				gMGC_Campus_KN = "일산 메가스터디학원"
				gMGC_Campus_EN = "ilsan"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "재수관"
				gMGC_Campus_Detail_Code = "CD0255"
				gMGC_Campus_DirSite ="/ilsan/nsu"
				gMGC_Campus_Title = cDev&"일산 메가스터디학원"
				gMGC_Campus_Theme = "red" 

			ElseIf InStr(gMGC_RequestPathInfo, "/yangju/jaehak") > 0 Then

				'양주 초.중.고
				gMGC_Campus_N = "yangju"
				gMGC_Campus_Code = "CD0258"
				gMGC_Campus_KN = "양주 메가스터디학원"
				gMGC_Campus_EN = "yangju"
				gMGC_Campus_Kind = "jaehak"
				gMGC_Campus_Kind_KN = "초중고관"
				gMGC_Campus_Detail_Code = "CD0259"
				gMGC_Campus_DirSite ="/yangju/jaehak"
				gMGC_Campus_Title = cDev&"양주 메가스터디학원 초중고관"
				gMGC_Campus_Theme = "blue"

			ElseIf InStr(gMGC_RequestPathInfo, "/unjung/jaehak") > 0 Then

				'운정 초.중.고
				gMGC_Campus_N = "unjung"
				gMGC_Campus_Code = "CD0260"
				gMGC_Campus_KN = "운정 메가스터디학원"
				gMGC_Campus_EN = "unjung"
				gMGC_Campus_Kind = "jaehak"
				gMGC_Campus_Kind_KN = "중고등관"
				gMGC_Campus_Detail_Code = "CD0261"
				gMGC_Campus_DirSite ="/unjung/jaehak"
				gMGC_Campus_Title = cDev&"운정 메가스터디학원 초중고관"
				gMGC_Campus_Theme = "blue"

			ElseIf InStr(gMGC_RequestPathInfo, "/songpa/nsu") > 0 Then

				'송파 재수관
				gMGC_Campus_N = "songpa"
				gMGC_Campus_Code = "CD0275"
				gMGC_Campus_KN = "송파 메가스터디학원"
				gMGC_Campus_EN = "songpa"
				gMGC_Campus_Kind = "nsu"
				gMGC_Campus_Kind_KN = "재수관"
				gMGC_Campus_Detail_Code = "CD0276"
				gMGC_Campus_DirSite ="/songpa/nsu"
				gMGC_Campus_Title = cDev&"송파 메가스터디학원"
				gMGC_Campus_Theme = "red"

			ElseIf InStr(gMGC_RequestPathInfo, "/ansung/") > 0 Then
				
				'안성기숙 재수정규반
				gMGC_Campus_N = "ansung"
				gMGC_Campus_Code = "CD0279"
				gMGC_Campus_KN = "안성 메가스터디 기숙학원"
				gMGC_Campus_EN = "ansung"
				gMGC_Campus_Kind = "gisuk"
				gMGC_Campus_Kind_KN = ""
				gMGC_Campus_Detail_Code = "CD0280"
				gMGC_Campus_DirSite ="/ansung"
				gMGC_Campus_Title = cDev&"안성 메가스터디 기숙학원"
				gMGC_Campus_Theme = "beige"

			Else
					' 그 외 
					gMGC_Campus_N = "campus"
					gMGC_Campus_Code = ""
					gMGC_Campus_KN = "메가스터디학원"
					gMGC_Campus_EN = "megastudy campus"
					gMGC_Campus_Kind = ""
					gMGC_Campus_Kind_KN = ""
					gMGC_Campus_Detail_Code = ""
					gMGC_Campus_DirSite ="/campus"
					gMGC_Campus_Title = cDev&"메가스터디학원"
					gMGC_Campus_Theme = ""

			End If

		Else	' 로컬 작업 아니면 다음 로직 탐
			
			Select Case gMGC_RequestServerName

				'Campus 
				Case "campus.megastudy.net", "campus1.megastudy.net", "campus2.megastudy.net", "campus5.megastudy.net", "amsdev.megastudy.net" ,"tcampus.megastudy.net", "mcampus.megastudy.net", "mcampus1.megastudy.net", "mcampus2.megastudy.net", "mcampus5.megastudy.net" , "tmcampus.megastudy.net" , "campus3.megastudy.net" , "mcampus3.megastudy.net"
		   
					gMGC_Campus_N = ""
					gMGC_Campus_Code = ""
					gMGC_Campus_KN = "메가스터디학원"
					gMGC_Campus_EN = "megastudy campus"
					gMGC_Campus_Kind = ""
					gMGC_Campus_Kind_KN = ""
					gMGC_Campus_Detail_Code = ""
					gMGC_Campus_DirSite ="/campus"
					gMGC_Campus_Title = cDev&"메가스터디학원"
					gMGC_Campus_Theme = ""
				
				'강남 재수관	http://gangnam.megastudy.net	http://mgangnam.megastudy.net
				Case "gangnam.megastudy.net", "mgangnam.megastudy.net", "tgangnam.megastudy.net", "tmgangnam.megastudy.net"

					gMGC_Campus_N = "gangnam"
					gMGC_Campus_Code = "CD0001"
					gMGC_Campus_KN = "강남 메가스터디학원"
					gMGC_Campus_EN = "gangnam"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "재수관"
					gMGC_Campus_Detail_Code = "CD0206"
					gMGC_Campus_DirSite ="/gangnam/nsu"
					gMGC_Campus_Title = cDev&"강남 메가스터디학원"
					gMGC_Campus_Theme = "red"
				
				'강남 중.고등관	http://gangnams.megastudy.net	http://mgangnams.megastudy.net
				Case "gangnams.megastudy.net", "mgangnams.megastudy.net", "tgangnams.megastudy.net", "tmgangnams.megastudy.net"
					
					gMGC_Campus_N = "gangnam"
					gMGC_Campus_Code = "CD0001"
					gMGC_Campus_KN = "강남 메가스터디학원"
					gMGC_Campus_EN = "gangnam"
					gMGC_Campus_Kind = "jaehak"
					gMGC_Campus_Kind_KN = "고등관"
					gMGC_Campus_Detail_Code = "CD0205"
					gMGC_Campus_DirSite ="/gangnam/jaehak"
					gMGC_Campus_Title = cDev&"강남 메가스터디학원 중고등관"
					gMGC_Campus_Theme = "blue"
					
				'강동 재수관	http://gangdong.megastudy.net	http://mgangdong.megastudy.net
				Case "gangdong.megastudy.net", "mgangdong.megastudy.net", "tgangdong.megastudy.net", "tmgangdong.megastudy.net"
					
					gMGC_Campus_N = "gangdong"
					gMGC_Campus_Code = "CD0198"
					gMGC_Campus_KN = "강동 메가스터디학원"
					gMGC_Campus_EN = "gangdong"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "재수관"
					gMGC_Campus_Detail_Code = "CD0240"
					gMGC_Campus_DirSite ="/gangdong/nsu"
					gMGC_Campus_Title = cDev&"강동 메가스터디학원"
					gMGC_Campus_Theme = "red"

				'강동 초.중.고	http://gangdongs.megastudy.net	http://mgangdongs.megastudy.net
				Case "gangdongs.megastudy.net", "mgangdongs.megastudy.net", "tgangdongs.megastudy.net", "tmgangdongs.megastudy.net"
					
					gMGC_Campus_N = "gangdong"
					gMGC_Campus_Code = "CD0198"
					gMGC_Campus_KN = "강동 메가스터디학원"
					gMGC_Campus_EN = "gangdong"
					gMGC_Campus_Kind = "jaehak"
					gMGC_Campus_Kind_KN = "초중고관"
					gMGC_Campus_Detail_Code = "CD0215"
					gMGC_Campus_DirSite ="/gangdong/jaehak"
					gMGC_Campus_Title = cDev&"강동 메가스터디학원"
					gMGC_Campus_Theme = "blue" 
				
				'강북 재수관	http://gangbuk.megastudy.net	http://mgangbuk.megastudy.net
				Case "gangbuk.megastudy.net", "mgangbuk.megastudy.net", "tgangbuk.megastudy.net", "tmgangbuk.megastudy.net"
				
					gMGC_Campus_N = "gangbuk"
					gMGC_Campus_Code = "CD0005"
					gMGC_Campus_KN = "강북 메가스터디학원"
					gMGC_Campus_EN = "gangbuk"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "재수관"
					gMGC_Campus_Detail_Code = "CD0210"
					gMGC_Campus_DirSite ="/gangbuk/nsu"
					gMGC_Campus_Title = cDev&"강북 메가스터디학원"
					gMGC_Campus_Theme = "red" 

				'강북 고등관	http://gangbuks.megastudy.net	http://mgangbuks.megastudy.net
				Case "gangbuks.megastudy.net", "mgangbuks.megastudy.net", "tgangbuks.megastudy.net", "tmgangbuks.megastudy.net"
				
					gMGC_Campus_N = "gangbuk"
					gMGC_Campus_Code = "CD0005"
					gMGC_Campus_KN = "강북 메가스터디학원"
					gMGC_Campus_EN = "gangbuk"
					gMGC_Campus_Kind = "jaehak"
					gMGC_Campus_Kind_KN = "고등 종합·단과학원"
					gMGC_Campus_Detail_Code = "CD0209"
					gMGC_Campus_DirSite ="/gangbuk/jaehak"
					gMGC_Campus_Title = cDev&"강북 메가스터디학원 고등관"
					gMGC_Campus_Theme = "blue"  

				'노량진 재수관	http://noryangjin.megastudy.net	http://mnoryangjin.megastudy.net
				Case "noryangjin.megastudy.net", "mnoryangjin.megastudy.net", "tnoryangjin.megastudy.net", "tmnoryangjin.megastudy.net"
					
					gMGC_Campus_N = "noryangjin"
					gMGC_Campus_Code = "CD0006"
					gMGC_Campus_KN = "노량진 메가스터디학원"
					gMGC_Campus_EN = "noryangjin"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "재수관"
					gMGC_Campus_Detail_Code = "CD0211"
					gMGC_Campus_DirSite ="/noryangjin/nsu"
					gMGC_Campus_Title = cDev&"노량진 메가스터디학원"
					gMGC_Campus_Theme = "red"   

				'노량진 단과	http://nrj.megastudy.net	http://mnrj.megastudy.net
				Case "nrj.megastudy.net", "mnrj.megastudy.net","tnrj.megastudy.net", "tmnrj.megastudy.net"
					
					gMGC_Campus_N = "noryangjin"
					gMGC_Campus_Code = "CD0006"
					gMGC_Campus_KN = "노량진 메가스터디학원"
					gMGC_Campus_EN = "noryangjin"
					gMGC_Campus_Kind = "danka"
					gMGC_Campus_Kind_KN = "단과"
					gMGC_Campus_Detail_Code = "CD0212"
					gMGC_Campus_DirSite ="/noryangjin/danka"
					gMGC_Campus_Title = cDev&"노량진 메가스터디학원 단과"
					gMGC_Campus_Theme = "blue"   
				
				'서초 재수관	http://seocho.megastudy.net	http://mseocho.megastudy.net
				Case "seocho.megastudy.net", "mseocho.megastudy.net", "tseocho.megastudy.net", "tmseocho.megastudy.net"
					
					gMGC_Campus_N = "seocho"
					gMGC_Campus_Code = "CD0004"
					gMGC_Campus_KN = "서초 메가스터디학원"
					gMGC_Campus_EN = "seocho"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "재수관"
					gMGC_Campus_Detail_Code = "CD0208"
					gMGC_Campus_DirSite ="/seocho/nsu"
					gMGC_Campus_Title = cDev&"서초 메가스터디학원"
					gMGC_Campus_Theme = "red" 

				'서초 고등관	http://seochos.megastudy.net	http://mseochos.megastudy.net
				Case "seochos.megastudy.net", "mseochos.megastudy.net", "tseochos.megastudy.net", "tmseochos.megastudy.net"
					
					gMGC_Campus_N = "seocho"
					gMGC_Campus_Code = "CD0004"
					gMGC_Campus_KN = "서초 메가스터디학원"
					gMGC_Campus_EN = "seocho"
					gMGC_Campus_Kind = "jaehak"
					gMGC_Campus_Kind_KN = "고등관"
					gMGC_Campus_Detail_Code = "CD0207"
					gMGC_Campus_DirSite ="/seocho/jaehak"
					gMGC_Campus_Title = cDev&"서초 메가스터디학원 고등관"
					gMGC_Campus_Theme = "blue"   
					
				'성북 재수관	http://seongbuk.megastudy.net	http://mseongbuk.megastudy.net
				Case "seongbuk.megastudy.net", "mseongbuk.megastudy.net", "tseongbuk.megastudy.net", "tmseongbuk.megastudy.net"
					
					gMGC_Campus_N = "seongbuk"
					gMGC_Campus_Code = "CD0098"
					gMGC_Campus_KN = "성북 메가스터디학원"
					gMGC_Campus_EN = "seongbuk"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "재수관"
					gMGC_Campus_Detail_Code = "CD0239"
					gMGC_Campus_DirSite ="/seongbuk/nsu"
					gMGC_Campus_Title = cDev&"성북 메가스터디학원"
					gMGC_Campus_Theme = "red"  
					
				'성북 초.중.고	http://seongbuks.megastudy.net	 http://mseongbuks.megastudy.net
				Case "seongbuks.megastudy.net", "mseongbuks.megastudy.net", "tseongbuks.megastudy.net", "tmseongbuks.megastudy.net"
					
					gMGC_Campus_N = "seongbuk"
					gMGC_Campus_Code = "CD0098"
					gMGC_Campus_KN = "성북 메가스터디학원"
					gMGC_Campus_EN = "seongbuk"
					gMGC_Campus_Kind = "jaehak"
					gMGC_Campus_Kind_KN = "초중고관"
					gMGC_Campus_Detail_Code = "CD0214"
					gMGC_Campus_DirSite ="/seongbuk/jaehak"
					gMGC_Campus_Title = cDev&"성북 메가스터디학원"
					gMGC_Campus_Theme = "blue"  
				
				'신촌 재수정규반	http://sinchon.megastudy.net	http://msinchon.megastudy.net
				Case "sinchon.megastudy.net", "msinchon.megastudy.net", "tsinchon.megastudy.net", "tmsinchon.megastudy.net"
					
					gMGC_Campus_N = "sinchon"
					gMGC_Campus_Code = "CD0036"
					gMGC_Campus_KN = "신촌 메가스터디학원"
					gMGC_Campus_EN = "sinchon"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "재수/단과"
					gMGC_Campus_Detail_Code = "CD0213"
					gMGC_Campus_DirSite ="/sinchon"
					gMGC_Campus_Title = cDev&"신촌 메가스터디학원"
					gMGC_Campus_Theme = "red"  

				'평촌 재수관	http://pyeongchon.megastudy.net	http://mpyeongchon.megastudy.net
				Case "pyeongchon.megastudy.net", "mpyeongchon.megastudy.net" , "tpyeongchon.megastudy.net", "tmpyeongchon.megastudy.net"
					
					gMGC_Campus_N = "pyeongchon"
					gMGC_Campus_Code = "CD0179"
					gMGC_Campus_KN = "평촌 메가스터디학원"
					gMGC_Campus_EN = "pyeongchon"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "재수관"
					gMGC_Campus_Detail_Code = "CD0217"
					gMGC_Campus_DirSite ="/pyeongchon/nsu"
					gMGC_Campus_Title = cDev&"평촌 메가스터디학원"
					gMGC_Campus_Theme = "red"  

				'평촌 고등관	http://pyeongchons.megastudy.net	http://mpyeongchons.megastudy.net
				Case "pyeongchons.megastudy.net", "mpyeongchons.megastudy.net","tpyeongchons.megastudy.net", "tmpyeongchons.megastudy.net"
					
					gMGC_Campus_N = "pyeongchon"
					gMGC_Campus_Code = "CD0179"
					gMGC_Campus_KN = "평촌 메가스터디학원"
					gMGC_Campus_EN = "pyeongchon"
					gMGC_Campus_Kind = "jaehak"
					gMGC_Campus_Kind_KN = "고등관"
					gMGC_Campus_Detail_Code = "CD0216"
					gMGC_Campus_DirSite ="/pyeongchon/jaehak"
					gMGC_Campus_Title = cDev&"평촌 메가스터디학원 고등관"
					gMGC_Campus_Theme = "blue"  

				'양지 본원 재수정규반	http://yangji.megastudy.net	http://myangji.megastudy.net
				Case "yangji.megastudy.net", "myangji.megastudy.net", "tyangji.megastudy.net", "tmyangji.megastudy.net"	
					
					gMGC_Campus_N = "yangji"
					gMGC_Campus_Code = "CD0178"
					gMGC_Campus_KN = "최상위권 전문관 러셀 기숙학원"
					gMGC_Campus_EN = "yangji"
					gMGC_Campus_Kind = "gisuk"
					gMGC_Campus_Kind_KN = ""
					gMGC_Campus_Detail_Code = "CD0220"
					gMGC_Campus_DirSite ="/yangji"
					gMGC_Campus_Title = cDev&"최상위권 전문관 러셀 기숙학원"
					gMGC_Campus_Theme = "beige" 

				'양지 신관 재수정규반	http://yangjim.megastudy.net	http://myangjim.megastudy.net
				Case "yangjim.megastudy.net", "myangjim.megastudy.net"	, "tyangjim.megastudy.net", "tmyangjim.megastudy.net" ,"seochob.megastudy.net", "mseochob.megastudy.net"	, "tseochob.megastudy.net", "tmseochob.megastudy.net"
					
					gMGC_Campus_N = "yangjim"
					gMGC_Campus_Code = "CD0039"
					gMGC_Campus_KN = "자연계 전문관 러셀 기숙학원"
					gMGC_Campus_EN = "yangjim"
					gMGC_Campus_Kind = "gisuk"
					gMGC_Campus_Kind_KN = ""
					gMGC_Campus_Detail_Code = "CD0243"
					gMGC_Campus_DirSite ="/yangjim"
					gMGC_Campus_Title = cDev&"자연계 전문관 러셀 기숙학원"
					gMGC_Campus_Theme = "beige"   'purple
				

				'러셀대치 http://campus.magastudy.net/russel/
				Case "russeldc.megastudy.net", "mrusseldc.megastudy.net"

					 gMGC_Campus_N = "russel"
					 gMGC_Campus_Code = "CD0001"
					 gMGC_Campus_KN = "메가스터디 러셀 대치"
					 gMGC_Campus_EN = "russel"
					 gMGC_Campus_Kind = "danka"
					 gMGC_Campus_Kind_KN = "단과"
					 gMGC_Campus_Detail_Code = "CD0241"
					 If gMGC_RequestServerName ="mrusseldc.megastudy.net" Then 
					 gMGC_Campus_DirSite ="/m_russel/"
					 Else 
					 gMGC_Campus_DirSite ="/russel/"
					 End If 
					 gMGC_Campus_Title = cDev&"러셀 대치 "
					 gMGC_Campus_Theme = ""

			   '러셀분당 http://campus.magastudy.net/russel_bundang/
			   Case "russelbd.megastudy.net", "mrusselbd.megastudy.net"

					gMGC_Campus_N = "russel_bundang"
					gMGC_Campus_Code = "CD0244"
					gMGC_Campus_KN = "메가스터디 러셀 분당"
					gMGC_Campus_EN = "russel_bundang"
					gMGC_Campus_Kind = "danka"
					gMGC_Campus_Kind_KN = "단과"
					gMGC_Campus_Detail_Code = "CD0244"
					gMGC_Campus_DirSite ="/russel_bundang"
					gMGC_Campus_Title = cDev&"러셀 분당"
					gMGC_Campus_Theme = ""

			   '러셀목동 http://campus.magastudy.net/russel_mokdong/
				Case "russelmd.megastudy.net", "mrusselmd.megastudy.net"

					gMGC_Campus_N = "russel_mokdong"
					gMGC_Campus_Code = "CD0245"
					gMGC_Campus_KN = "메가스터디 러셀 목동"
					gMGC_Campus_EN = "russel_mokdong"
					gMGC_Campus_Kind = "danka"
					gMGC_Campus_Kind_KN = "단과"
					gMGC_Campus_Detail_Code = "CD0245"
					gMGC_Campus_DirSite ="/russel_mokdong"
					gMGC_Campus_Title = cDev&"러셀 목동"
					gMGC_Campus_Theme = ""
				
				Case "russelct.megastudy.net", "russelct.megastudy.net"

					gMGC_Campus_N = "russel_ct"
					gMGC_Campus_Code = "CD0246"
					gMGC_Campus_KN = "메가스터디 러셀 센텀"
					gMGC_Campus_EN = "russel_ct"
					gMGC_Campus_Kind = "danka"
					gMGC_Campus_Kind_KN = "단과"
					gMGC_Campus_Detail_Code = "CD0246"
					gMGC_Campus_DirSite ="/russel_ct"
					gMGC_Campus_Title = cDev&"러셀 센텀"
					gMGC_Campus_Theme = ""

				Case "russelgn.megastudy.net", "russelgn.megastudy.net"

					gMGC_Campus_N = "russel_gn"
					gMGC_Campus_Code = "CD0247"
					gMGC_Campus_KN = "메가스터디 러셀 강남"
					gMGC_Campus_EN = "russel_gn"
					gMGC_Campus_Kind = "danka"
					gMGC_Campus_Kind_KN = "단과"
					gMGC_Campus_Detail_Code = "CD0247"
					gMGC_Campus_DirSite ="/russel_gn"
					gMGC_Campus_Title = cDev&"러셀 강남"
					gMGC_Campus_Theme = ""
				
				Case "russelpc.megastudy.net", "russelpc.megastudy.net"

					gMGC_Campus_N = "russel_pc"
					gMGC_Campus_Code = "CD0248"
					gMGC_Campus_KN = "메가스터디 러셀 평촌"
					gMGC_Campus_EN = "russel_pc"
					gMGC_Campus_Kind = "danka"
					gMGC_Campus_Kind_KN = "단과"
					gMGC_Campus_Detail_Code = "CD0248"
					gMGC_Campus_DirSite ="/russel_pc"
					gMGC_Campus_Title = cDev&"러셀 평촌"
					gMGC_Campus_Theme = ""

				
				Case "russelyt.megastudy.net", "russelyt.megastudy.net"

					gMGC_Campus_N = "russel_yt"
					gMGC_Campus_Code = "CD0249"
					gMGC_Campus_KN = "메가스터디 러셀 영통"
					gMGC_Campus_EN = "russel_yt"
					gMGC_Campus_Kind = "danka"
					gMGC_Campus_Kind_KN = "단과"
					gMGC_Campus_Detail_Code = "CD0249"
					gMGC_Campus_DirSite ="/russel_yt"
					gMGC_Campus_Title = cDev&"러셀 영통"
					gMGC_Campus_Theme = ""

				
				Case "russelbc.megastudy.net", "russelbc.megastudy.net"

					gMGC_Campus_N = "russel_bc"
					gMGC_Campus_Code = "CD0250"
					gMGC_Campus_KN = "메가스터디 러셀 부천"
					gMGC_Campus_EN = "russel_bc"
					gMGC_Campus_Kind = "danka"
					gMGC_Campus_Kind_KN = "단과"
					gMGC_Campus_Detail_Code = "CD0250"
					gMGC_Campus_DirSite ="/russel_bc"
					gMGC_Campus_Title = cDev&"러셀 부천"
					gMGC_Campus_Theme = ""
				
				Case "russel.megastudy.net"

					gMGC_Campus_N = "intro"
					gMGC_Campus_Code = "RUSSELINTRO"
					gMGC_Campus_KN = "메가스터디 러셀"
					gMGC_Campus_EN = "intro"
					gMGC_Campus_Kind = "danka"
					gMGC_Campus_Kind_KN = "단과"
					gMGC_Campus_Detail_Code = "RUSSELINTRO"
					gMGC_Campus_DirSite ="/intro"
					gMGC_Campus_Title = cDev&"러셀 INTRO"
					gMGC_Campus_Theme = ""
				
				'러셀중계 http://campus.magastudy.net/russel_jg/
				Case "russeljg.megastudy.net", "mrusseljg.megastudy.net"
					gMGC_Campus_N = "russel_jg"
					gMGC_Campus_Code = "CD0257"
					gMGC_Campus_KN = "메가스터디 러셀 중계"
					gMGC_Campus_EN = "russel_jg"
					gMGC_Campus_Kind = "danka"
					gMGC_Campus_Kind_KN = "단과"
					gMGC_Campus_Detail_Code = "CD0257"
					gMGC_Campus_DirSite ="/russel_jg/"
					gMGC_Campus_Title = cDev&"러셀 중계"
					gMGC_Campus_Theme = ""

				'러셀 기숙
				Case "russelcampus.megastudy.net", "mrusselcampus.megastudy.net"
					gMGC_Campus_N = "russel_yj"
					gMGC_Campus_Code = "CD0258"
					gMGC_Campus_KN = "메가스터디 러셀 기숙"
					gMGC_Campus_EN = "russel_yj"
					gMGC_Campus_Kind = "gisuk"
					gMGC_Campus_Kind_KN = "기숙"
					gMGC_Campus_Detail_Code = "CD0258"
					gMGC_Campus_DirSite ="/russel_yj/"
					gMGC_Campus_Title = cDev&"러셀 기숙"
					gMGC_Campus_Theme = ""


				'부천 재수정규반	http://bucheon.megastudy.net	http://mbucheon.megastudy.net
				Case "bucheon.megastudy.net", "mbucheon.megastudy.net", "tbucheon.megastudy.net", "tmbucheon.megastudy.net"
				
					gMGC_Campus_N = "bucheon"
					gMGC_Campus_Code = "CD0250"
					gMGC_Campus_KN = "부천 메가스터디학원"
					gMGC_Campus_EN = "bucheon"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "재수관"
					gMGC_Campus_Detail_Code = "CD0251"
					gMGC_Campus_DirSite ="/bucheon/nsu"
					gMGC_Campus_Title = cDev&"부천 메가스터디학원"
					gMGC_Campus_Theme = "red" 

				'분당 재수정규반	http://bundang.megastudy.net	http://mbundang.megastudy.net
				Case "bundang.megastudy.net", "mbundang.megastudy.net", "tbundang.megastudy.net", "tmbundang.megastudy.net"
				
					gMGC_Campus_N = "bundang"
					gMGC_Campus_Code = "CD0252"
					gMGC_Campus_KN = "분당 메가스터디학원"
					gMGC_Campus_EN = "bundang"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "재수관"
					gMGC_Campus_Detail_Code = "CD0253"
					gMGC_Campus_DirSite ="/bundang/nsu"
					gMGC_Campus_Title = cDev&"분당 메가스터디학원"
					gMGC_Campus_Theme = "red" 
			

					'일산 재수정규반	http://ilsan.megastudy.net	http://milsan.megastudy.net
				Case "ilsan.megastudy.net", "milsan.megastudy.net", "tilsan.megastudy.net", "tmilsan.megastudy.net"
				
					gMGC_Campus_N = "ilsan"
					gMGC_Campus_Code = "CD0254"
					gMGC_Campus_KN = "일산 메가스터디학원"
					gMGC_Campus_EN = "ilsan"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "재수관"
					gMGC_Campus_Detail_Code = "CD0255"
					gMGC_Campus_DirSite ="/ilsan/nsu"
					gMGC_Campus_Title = cDev&"일산 메가스터디학원"
					gMGC_Campus_Theme = "red" 

				'양주 초.중.고	http://yangju.megastudy.net	 http://myangju.megastudy.net
				Case "yangju.megastudy.net", "myangju.megastudy.net", "tyangju.megastudy.net", "tmyangju.megastudy.net"

					gMGC_Campus_N = "yangju"
					gMGC_Campus_Code = "CD0258"
					gMGC_Campus_KN = "양주 메가스터디학원"
					gMGC_Campus_EN = "yangju"
					gMGC_Campus_Kind = "jaehak"
					gMGC_Campus_Kind_KN = "초중고관"
					gMGC_Campus_Detail_Code = "CD0259"
					gMGC_Campus_DirSite ="/yangju/jaehak"
					gMGC_Campus_Title = cDev&"양주 메가스터디학원"
					gMGC_Campus_Theme = "blue"

				'운정 초.중.고	http://unjung.megastudy.net	 http://munjung.megastudy.net
				Case "unjung.megastudy.net", "munjung.megastudy.net", "tunjung.megastudy.net", "tmunjung.megastudy.net"

					gMGC_Campus_N = "unjung"
					gMGC_Campus_Code = "CD0260"
					gMGC_Campus_KN = "운정 메가스터디학원"
					gMGC_Campus_EN = "unjung"
					gMGC_Campus_Kind = "jaehak"
					gMGC_Campus_Kind_KN = "중고등관"
					gMGC_Campus_Detail_Code = "CD0261"
					gMGC_Campus_DirSite ="/unjung/jaehak"
					gMGC_Campus_Title = cDev&"운정 메가스터디학원"
					gMGC_Campus_Theme = "blue"

				'송파 재수관	http://songpa.megastudy.net	http://msongpa.megastudy.net
				Case "songpa.megastudy.net", "msongpa.megastudy.net", "tsongpa.megastudy.net", "tmsongpa.megastudy.net"
					
					gMGC_Campus_N = "songpa"
					gMGC_Campus_Code = "CD0275"
					gMGC_Campus_KN = "송파 메가스터디학원"
					gMGC_Campus_EN = "songpa"
					gMGC_Campus_Kind = "nsu"
					gMGC_Campus_Kind_KN = "재수관"
					gMGC_Campus_Detail_Code = "CD0276"
					gMGC_Campus_DirSite ="/songpa/nsu"
					gMGC_Campus_Title = cDev&"송파 메가스터디학원"
					gMGC_Campus_Theme = "red"

				'안성기숙 재수정규반	http://ansung.megastudy.net	http://mansung.megastudy.net
				Case "ansung.megastudy.net", "mansung.megastudy.net", "tansung.megastudy.net", "tmansung.megastudy.net"
					
					gMGC_Campus_N = "ansung"
					gMGC_Campus_Code = "CD0279"
					gMGC_Campus_KN = "안성 메가스터디 기숙학원"
					gMGC_Campus_EN = "ansung"
					gMGC_Campus_Kind = "gisuk"
					gMGC_Campus_Kind_KN = ""
					gMGC_Campus_Detail_Code = "CD0280"
					gMGC_Campus_DirSite ="/ansung"
					gMGC_Campus_Title = cDev&"안성 메가스터디 기숙학원"
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
					gMGC_Campus_Title = cDev&"메가스터디학원"
					gMGC_Campus_Theme = ""    
				
			End Select

		End If	' 로컬 작업 구분값 끝

		'사이트 변수 설정 E

		returnVal = ""
		
		'소스차이로 인한 문제발생 제거 하기 위한 임시 로직 
		If ValueCheck(Request.Cookies("_CAMPUS_CODE_") ) Then 
			TmpCheckValue = True 
		Else 
			TmpCheckValue = False  
		End If 

		TmpCheckValue = False  

		If TmpCheckValue Then 
			'도메인 쿠키 값 리턴 S
			Select Case UCase(CStr(VarsName))
				Case "_CAMPUS_" '학원 변수
					returnVal = Request.Cookies("_CAMPUS_")
				Case "_CAMPUS_CODE_" '학원 구코드 변수
					returnVal = Request.Cookies("_CAMPUS_CODE_")
				Case "_CAMPUS_NAME_" '학원이름 변수
					returnVal = Request.Cookies("_CAMPUS_NAME_")
				Case "_CAMPUS_ENG_NAME_" '학원 영문명
					returnVal = Request.Cookies("_CAMPUS_ENG_NAME_")
				Case "_CAMPUS_KIND_" '학원 단과,재수,기숙 변수
					returnVal = Request.Cookies("_CAMPUS_KIND_")
					'모바일 returnVal = Request.Cookies("_CAMPUS_CATE_")
				Case "_CAMPUS_DETAIL_CODE_" '학원 신코드 변수			
					returnVal = Request.Cookies("_CAMPUS_DETAIL_CODE_")
					'모바일 returnVal = Request.Cookies("_CAMPUS_CODE_")
				Case "_CAMPUS_DIRSITE_" '학원 디렉토리 변수			
					returnVal = Request.Cookies("_CAMPUS_DIRSITE_")
				Case "_CAMPUS_KIND_NAME_" '학원 카테고리 이름 변수 
					returnVal = Request.Cookies("_CAMPUS_CATE_NAME_")
				Case "_CAMPUS_THEME_" '모바일 학원 테마 
					returnVal = Request.Cookies("_CAMPUS_THEME_")
				Case Else '그외 변수
					returnVal = ""
			End Select
			'도메인 쿠키 값 리턴 E
		Else 
			'도메인 변수 값 리턴 S
			Select Case UCase(CStr(VarsName))
				Case "_CAMPUS_" '학원 변수
					returnVal = gMGC_Campus_N
				Case "_CAMPUS_CODE_" '학원 구코드 변수
					returnVal = gMGC_Campus_Code
				Case "_CAMPUS_NAME_" '학원이름 변수
					returnVal = gMGC_Campus_KN
				Case "_CAMPUS_ENG_NAME_" '학원 영문명
					returnVal = gMGC_Campus_EN	
				Case "_CAMPUS_KIND_" '학원 단과,재수,기숙 변수
					returnVal = gMGC_Campus_Kind
				Case "_CAMPUS_DETAIL_CODE_" '학원 신코드 변수			
					returnVal = gMGC_Campus_Detail_Code
				Case "_CAMPUS_DIRSITE_" '학원 디렉토리 변수			
					returnVal = gMGC_Campus_DirSite
				Case "_CAMPUS_KIND_NAME_" '학원 카테고리 이름 변수 
					returnVal = gMGC_Campus_Kind_KN	
				Case "_CAMPUS_THEME_" '모바일 학원 테마 
					returnVal = gMGC_Campus_Theme
				Case "_CAMPUS_TITLE_" '모바일 학원 테마 
					returnVal = gMGC_Campus_Title
				Case Else '그외 변수
					returnVal = ""
			End Select
			'도메인 변수 값 리턴 E
		End If 

		GetCampusVarsSiteMGC =  returnVal
	End Function 
	'학원별 변수 가져오기 함수 E

	' 학원별 사이트주소 가져오기  S
	Function GetCampusDetailUrlSiteMGC(sCampusCodeDetail,isDirCheck)
		returnURLSite = ""
		isCheckDev = False 

		If Request.ServerVariables("LOCAL_ADDR") = "10.1.3.10" Then  '개발서버 체크
			isCheckDev = True 
		Else
			isCheckDev = False 
		End If

		If Request.ServerVariables("LOCAL_ADDR") = "::1" Then	' 로컬에서 작업을 위해서 추가 (2020.03.10)
			Select Case UCase(CStr(sCampusCodeDetail))
				Case "INTRO" '인트로페이지로
					If isDirCheck Then 
						returnURLSite = "/campus"
					End If 
				Case "CD0205" '강남고등관
					If isDirCheck Then 
						returnURLSite = "/gangnam/jaehak"
					End If 
				Case "CD0206" '강남재수관
					If isDirCheck Then 
						returnURLSite = "/gangnam/nsu"
					End If 
				Case "CD0207" '서초고등관
					If isDirCheck Then 
						returnURLSite = "/seocho/jaehak"
					End If 
				Case "CD0208" '서초재수관
					If isDirCheck Then 
						returnURLSite = "/seocho/nsu"
					End If 
				Case "CD0209" '강북고등관
					If isDirCheck Then 
						returnURLSite = "/gangbuk/jaehak"
					End If 
				Case "CD0210" '강북재수관
					If isDirCheck Then 
						returnURLSite = "/gangbuk/nsu"
					End If 
				Case "CD0211" '노량진재수관
					If isDirCheck Then 
						returnURLSite = "/noryangjin/nsu"
					End If 
				Case "CD0212" '노량진단과
					If isDirCheck Then 
						returnURLSite = "/noryangjin/danka"
					End If 
				Case "CD0213" '신촌
					If isDirCheck Then 
						returnURLSite = "/sinchon"
					End If 
				Case "CD0214" '성북고등관
					If isDirCheck Then 
						returnURLSite = "/seongbuk/jaehak"
					End If 
				Case "CD0239" '성북재수관
					If isDirCheck Then 
						returnURLSite = "/seongbuk/nsu"
					End If 
				Case "CD0215" '강동고등관
					If isDirCheck Then 
						returnURLSite = "/gangdong/jaehak"
					End If 

				Case "CD0240" '강동재수관
					If isDirCheck Then 
						returnURLSite = "/gangdong/nsu"
					End If 
				Case "CD0216" '평촌고등관
					If isDirCheck Then 
						returnURLSite = "/pyeongchon/jaehak"
					End If 
				Case "CD0217" '평촌재수관
					If isDirCheck Then 
						returnURLSite = "/pyeongchon/nsu"
					End If 
				Case "CD0218" '광주
					If isDirCheck Then 
						returnURLSite = returnURLSite & ""
					End If 
				Case "CD0219" '남양주
					If isCheckDev Then 
						returnURLSite = ""
					Else 
						returnURLSite = ""
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & ""
					End If 
				Case "CD0220" '양지본관
					If isDirCheck Then 
						returnURLSite = "/yangji"
					End If 
				Case "CD0243" '양지신관
					If isDirCheck Then 
						returnURLSite = "/yangjim"
					End If 
				Case "CD0221" '수원 -- 쓸일없음
					If isCheckDev Then 
						returnURLSite = ""
					Else 
						returnURLSite = ""
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & ""
					End If 
				Case "CD0241" '러셀 대치
					  If isDirCheck Then 
						returnURLSite = "/russel"
					  End If 

				Case "CD0244" '러셀 분당
					  If isDirCheck Then 
						 returnURLSite = "/russel_bundang"
					  End If 

				Case "CD0245" '러셀 목동
						If isDirCheck Then 
						  returnURLSite = "/russel_mokdong"
						End If
				
				 Case "CD0246" '러셀 센텀
						If isDirCheck Then 
						  returnURLSite = "/russel_ct"
						End If

				Case "CD0247" '러셀 강남
						If isDirCheck Then 
						  returnURLSite = "/russel_gn"
						End If
				
				Case "CD0248" '러셀 편촌
						If isDirCheck Then 
						  returnURLSite = "/russel_pc"
						End If
				
				Case "CD0249" '러셀 영통
						If isDirCheck Then 
						  returnURLSite = "/russel_yt"
						End If

				Case "CD0250" '러셀 부천
						If isDirCheck Then 
						  returnURLSite = "/russel_bc"
						End If
				Case "RUSSELINTRO" '러셀 INTRO
						If isDirCheck Then 
						  returnURLSite = "/intro"
						End If


				Case "CD0251" '부천 재수관
					If isDirCheck Then 
						returnURLSite = "/bucheon/nsu"
					End If 

				Case "CD0253" '분당 재수관
					If isDirCheck Then 
						returnURLSite = "/bundang/nsu"
					End If 


				Case "CD0255" '일산 재수관
					If isDirCheck Then 
						returnURLSite = "/ilsan/nsu"
					End If 


				Case "CD0259" '양주고등관
					If isDirCheck Then
						returnURLSite = "/yangju/jaehak"
					End If

				Case "CD0261" '운정고등관
					If isDirCheck Then
						returnURLSite = "/unjung/jaehak"
					End If

				Case "CD0276" '송파재수관
					If isDirCheck Then 
						returnURLSite = "/songpa/nsu"
					End If 

				Case "CD0280" '안성기숙
					If isDirCheck Then 
						returnURLSite = "/ansung"
					End If 

				Case Else '그 이외의경우 
					If isDirCheck Then 
						returnURLSite = "/campus"
					End If 

			End Select

		Else

			Select Case UCase(CStr(sCampusCodeDetail))
				Case "INTRO" '인트로페이지로
					If isCheckDev Then 
						returnURLSite = "http://tcampus.megastudy.net"
					Else 
						returnURLSite = "http://campus.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/campus"
					End If 
				Case "CD0205" '강남고등관
					If isCheckDev Then 
						returnURLSite = "http://tgangnam.megastudy.net"
					Else 
						returnURLSite = "http://gangnam.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/gangnam/jaehak"
					End If 
				Case "CD0206" '강남재수관
					If isCheckDev Then 
						returnURLSite = "http://tgangnam.megastudy.net"
					Else 
						returnURLSite = "http://gangnam.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/gangnam/nsu"
					End If 
				Case "CD0207" '서초고등관
					If isCheckDev Then 
						returnURLSite = "http://tseocho.megastudy.net"
					Else 
						returnURLSite = "http://seocho.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/seocho/jaehak"
					End If 
				Case "CD0208" '서초재수관
					If isCheckDev Then 
						returnURLSite = "http://tseocho.megastudy.net"
					Else 
						returnURLSite = "http://seocho.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/seocho/nsu"
					End If 
				Case "CD0209" '강북고등관
					If isCheckDev Then 
						returnURLSite = "http://tgangbuks.megastudy.net"
					Else 
						returnURLSite = "http://gangbuks.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/gangbuk/jaehak"
					End If 
				Case "CD0210" '강북재수관
					If isCheckDev Then 
						returnURLSite = "http://tgangbuk.megastudy.net"
					Else 
						returnURLSite = "http://gangbuk.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/gangbuk/nsu"
					End If 
				Case "CD0211" '노량진재수관
					If isCheckDev Then 
						returnURLSite = "http://tnoryangjin.megastudy.net"
					Else 
						returnURLSite = "http://noryangjin.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/noryangjin/nsu"
					End If 
				Case "CD0212" '노량진단과
					If isCheckDev Then 
						returnURLSite = "http://tnoryangjin.megastudy.net"
					Else 
						returnURLSite = "http://noryangjin.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/noryangjin/danka"
					End If 
				Case "CD0213" '신촌
					If isCheckDev Then 
						returnURLSite = "http://tsinchon.megastudy.net"
					Else 
						returnURLSite = "http://sinchon.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/sinchon"
					End If 
				Case "CD0214" '성북고등관
					If isCheckDev Then 
						returnURLSite = "http://tseongbuks.megastudy.net"
					Else 
						returnURLSite = "http://seongbuks.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/seongbuk/jaehak"
					End If 
				Case "CD0239" '성북재수관
					If isCheckDev Then 
						returnURLSite = "http://tseongbuk.megastudy.net"
					Else 
						returnURLSite = "http://seongbuk.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/seongbuk/nsu"
					End If 
				Case "CD0215" '강동고등관
					If isCheckDev Then 
						returnURLSite = "http://tgangdongs.megastudy.net"
					Else 
						returnURLSite = "http://gangdongs.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/gangdong/jaehak"
					End If 

				Case "CD0240" '강동재수관
					If isCheckDev Then 
						returnURLSite = "http://tgangdong.megastudy.net"
					Else 
						returnURLSite = "http://gangdong.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/gangdong/nsu"
					End If 
				Case "CD0216" '평촌고등관
					If isCheckDev Then 
						returnURLSite = "http://tpyeongchon.megastudy.net"
					Else 
						returnURLSite = "http://pyeongchon.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/pyeongchon/jaehak"
					End If 
				Case "CD0217" '평촌재수관
					If isCheckDev Then 
						returnURLSite = "http://tpyeongchon.megastudy.net"
					Else 
						returnURLSite = "http://pyeongchon.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/pyeongchon/nsu"
					End If 
				Case "CD0218" '광주
					If isCheckDev Then 
						returnURLSite = ""
					Else 
						returnURLSite = ""
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & ""
					End If 
				Case "CD0219" '남양주
					If isCheckDev Then 
						returnURLSite = ""
					Else 
						returnURLSite = ""
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & ""
					End If 
				Case "CD0220" '양지본관
					If isCheckDev Then 
						returnURLSite = "http://tyangji.megastudy.net"
					Else 
						returnURLSite = "http://yangji.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/yangji"
					End If 
				Case "CD0243" '양지신관
					If isCheckDev Then 
						'returnURLSite = "http://tyangjim.megastudy.net"
						returnURLSite = "http://tseochob.megastudy.net"
					Else 
						'returnURLSite = "http://yangjim.megastudy.net"
						returnURLSite = "http://seochob.megastudy.net"
					End If 

					gMGC_RequestServerName_SUB = LCase(Request.ServerVariables("SERVER_NAME"))  '도메인 이름

					If gMGC_RequestServerName_SUB = "seochob.megastudy.net" Or gMGC_RequestServerName_SUB = "mseochob.megastudy.net" Or gMGC_RequestServerName_SUB ="tseochob.megastudy.net" Or gMGC_RequestServerName_SUB = "tmseochob.megastudy.net" Then 
						returnURLSite = "http://"&gMGC_RequestServerName_SUB
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/yangjim"
					End If 
				Case "CD0221" '수원 -- 쓸일없음
					If isCheckDev Then 
						returnURLSite = ""
					Else 
						returnURLSite = ""
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & ""
					End If 
				Case "CD0241" '러셀 대치
					  If isCheckDev Then 
						 returnURLSite = "http://russeldc.megastudy.net"
					  Else 
						 returnURLSite = "http://russeldc.megastudy.net"
					  End If 

					  If isDirCheck Then 
						returnURLSite = returnURLSite & "/russel"
					  End If 

				Case "CD0244" '러셀 분당
					  If isCheckDev Then 
						 returnURLSite = "http://russelbd.megastudy.net"
					  Else 
						 returnURLSite = "http://russelbd.megastudy.net"
					  End If 

					  If isDirCheck Then 
						 returnURLSite = returnURLSite & "/russel_bundang"
					  End If 

				Case "CD0245" '러셀 목동
					   If isCheckDev Then 
						 returnURLSite = "http://russelmd.megastudy.net"
					   Else 
						 returnURLSite = "http://russelmd.megastudy.net"
					   End If 

						If isDirCheck Then 
						  returnURLSite = returnURLSite & "/russel_mokdong"
						End If
				
				 Case "CD0246" '러셀 센텀
					   If isCheckDev Then 
						 returnURLSite = "http://russelct.megastudy.net"
					   Else 
						 returnURLSite = "http://russelct.megastudy.net"
					   End If 

						If isDirCheck Then 
						  returnURLSite = returnURLSite & "/russel_ct"
						End If

				Case "CD0247" '러셀 강남
					   If isCheckDev Then 
						 returnURLSite = "http://russelgn.megastudy.net"
					   Else 
						 returnURLSite = "http://russelgn.megastudy.net"
					   End If 

						If isDirCheck Then 
						  returnURLSite = returnURLSite & "/russel_gn"
						End If
				
				Case "CD0248" '러셀 편촌
					   If isCheckDev Then 
						 returnURLSite = "http://russelpc.megastudy.net"
					   Else 
						 returnURLSite = "http://russelpc.megastudy.net"
					   End If 

						If isDirCheck Then 
						  returnURLSite = returnURLSite & "/russel_pc"
						End If
				
				Case "CD0249" '러셀 영통
					   If isCheckDev Then 
						 returnURLSite = "http://russelyt.megastudy.net"
					   Else 
						 returnURLSite = "http://russelyt.megastudy.net"
					   End If 

						If isDirCheck Then 
						  returnURLSite = returnURLSite & "/russel_yt"
						End If

				Case "CD0250" '러셀 부천
					   If isCheckDev Then 
						 returnURLSite = "http://russelbc.megastudy.net"
					   Else 
						 returnURLSite = "http://russelbc.megastudy.net"
					   End If 

						If isDirCheck Then 
						  returnURLSite = returnURLSite & "/russel_bc"
						End If
				Case "RUSSELINTRO" '러셀 INTRO
					   If isCheckDev Then 
						 returnURLSite = "http://russel.megastudy.net"
					   Else 
						 returnURLSite = "http://russel.megastudy.net"
					   End If 

						If isDirCheck Then 
						  returnURLSite = returnURLSite & "/intro"
						End If


				Case "CD0251" '부천 재수관
					If isCheckDev Then 
						returnURLSite = "http://tbucheon.megastudy.net"
					Else 
						returnURLSite = "http://bucheon.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/bucheon/nsu"
					End If 

				Case "CD0253" '분당 재수관
					If isCheckDev Then 
						returnURLSite = "http://tbundang.megastudy.net"
					Else 
						returnURLSite = "http://bundang.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/bundang/nsu"
					End If 


				Case "CD0255" '일산 재수관
					If isCheckDev Then 
						returnURLSite = "http://tilsan.megastudy.net"
					Else 
						returnURLSite = "http://ilsan.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/ilsan/nsu"
					End If 


				Case "CD0259" '양주고등관
					If isCheckDev Then
						returnURLSite = "http://tyangju.megastudy.net"
					Else
						returnURLSite = "http://yangju.megastudy.net"
					End If

					If isDirCheck Then
						returnURLSite = returnURLSite & "/yangju/jaehak"
					End If

				Case "CD0261" '운정고등관
					If isCheckDev Then
						returnURLSite = "http://tunjung.megastudy.net"
					Else
						returnURLSite = "http://unjung.megastudy.net"
					End If

					If isDirCheck Then
						returnURLSite = returnURLSite & "/unjung/jaehak"
					End If

				Case "CD0276" '송파재수관
					If isCheckDev Then 
						returnURLSite = "http://tsongpa.megastudy.net"
					Else 
						returnURLSite = "http://songpa.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/songpa/nsu"
					End If 

				Case "CD0280" '안성기숙
					If isCheckDev Then 
						returnURLSite = "http://tansung.megastudy.net"
					Else 
						returnURLSite = "http://ansung.megastudy.net"
					End If 

					If isDirCheck Then 
						returnURLSite = returnURLSite & "/ansung"
					End If 

				Case Else '그 이외의경우 
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

	' 학원별 사이트주소 가져오기  E

	' 학원별 메인 index.asp 페이지 여부 체크  S
	Function GetCampusIsIndexMGC()
		returnVal = False 
		' 페이지 링크 정보
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

		'Index 여부 체크 
		If gMGC_SiteMenu_1th = "index.asp"  Then 
			returnVal = True  
		End If 

		GetCampusIsIndexMGC = returnVal
	End Function 
	' 학원별 메인 index.asp 페이지 여부 체크 E


	' 테스트 운영 체크   S
	Function GetCampusBaseUrlMGC()
		returnVal = "" 
		If Request.ServerVariables("LOCAL_ADDR") = "10.1.3.10" Then  '개발서버 체크
				returnVal = "http://tcampus.megastudy.net"
		Else
				returnVal = "http://campus.megastudy.net"
		End If	

		GetCampusBaseUrlMGC = returnVal
	End Function 
	' 테스트 운영 체크 E



	' campus 가 아닌 각 도메인 사이트 여부 체크   S
	Function GetCampusIsDMMGC()
		gMGC_Sub_RequestServerName = LCase(Request.ServerVariables("SERVER_NAME"))  '도메인 이름
		gMGC_Sub_ServerName = "http://" & RequestServerName	
		returnVal = False 
		If gMGC_Sub_RequestServerName="campus.megastudy.net" Or gMGC_Sub_RequestServerName="campus1.megastudy.net" Or gMGC_Sub_RequestServerName="campus2.megastudy.net" Or gMGC_Sub_RequestServerName="campus3.megastudy.net" Or gMGC_Sub_RequestServerName="campus5.megastudy.net" Or gMGC_Sub_RequestServerName="amsdev.megastudy.net" Or gMGC_Sub_RequestServerName="tcampus.megastudy.net" Or gMGC_Sub_RequestServerName="mcampus.megastudy.net" Or gMGC_Sub_RequestServerName="mcampus1.megastudy.net" Or gMGC_Sub_RequestServerName="mcampus2.megastudy.net" Or gMGC_Sub_RequestServerName="mcampus3.megastudy.net" Or gMGC_Sub_RequestServerName="mcampus5.megastudy.net" Or gMGC_Sub_RequestServerName="tmcampus.megastudy.net" Then 		
			returnVal = False
		Else		
			returnVal = True 
		End If 

		GetCampusIsDMMGC = returnVal
	End Function 
	' campus 가 아닌 각 도메인 사이트 여부 체크 E

	'페이지 이동 S

	If VARS("clear") = "Y" And GetCampusIsIndexMGC() Then  ' index.asp 페이지이고 clear=Y 파라미터이면  
			
			If VARS("clear") = "Y" And ( GetCampusIsDMMGC() ) Then ' Y 이고 다른 메인 캠퍼스 주소이면 이동 
				Response.redirect  GetCampusBaseUrlMGC()
				Response.End
			End If

	ElseIf ValueCheck( VARS("refer") ) And GetCampusIsIndexMGC() Then  ' index.asp 페이지이고 refer 파라미터가 있으면 해당 페이지로 이동   
				
		If GetCampusIsDMMGC() Then 		 ' 다른 메인 캠퍼스 주소이면 이동 
			Response.redirect  GetCampusBaseUrlMGC() & VARS("refer")
			Response.End
		End If 
		
	ElseIf GetCampusIsDMMGC() And GetCampusIsIndexMGC() Then  ' campus 가 아닌 각 도메인 사이트이면  index.asp 페이지이면 각 도메인학원으로 이동 
		'해당 각 학원사이트 이동
		If ValueCheck(GetCampusVarsSiteMGC("_CAMPUS_DIRSITE_")) Then
			If LCase(Request.ServerVariables("SERVER_NAME")) = "localhost" Then	
				' 로컬일때는 리다이렉트 안시킴(2020.03.10)
			Else
				Response.Redirect GetCampusVarsSiteMGC("_CAMPUS_DIRSITE_")
				Response.End
			End If
		End If 
	
	ElseIf (Not GetCampusIsDMMGC() ) And GetCampusIsIndexMGC() Then  'campus 사이트 이고 index.asp 페이지 이면 

	Else 

	End If 

	gMGC_Campus_Title = GetCampusVarsSiteMGC("_CAMPUS_TITLE_")

	'페이지 이동 E
	
	'도메인 분기 처리 E
%>
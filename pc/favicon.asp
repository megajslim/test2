<%
'##################################################################
' 파일설명 : 페이지마다 공통으로 사용되는 파일의
'            업데이트 관리를 위해 사용
' 생성     : 2015/11/13 - 주지호
' 사용법   : 수정사항이 생길 경우 ? 뒤의 값을 변경해,
'            사용자의 웹브라우져에게 새로운 파일인 것 처럼 보이게 한다
'##################################################################
%>
<% 
	'네이버 메타 체크 가져오기 함수 가져오기 함수 S 
	Function GetCampusNaverCheckMeta()
		
		isNaverCheck_CampusDetailCode = "0" 
		detailNaverCheck_CampusDetailCode = "9d6d7d0820f3efbf43bd3fd3ba73ca6f7449e3b5"

		NaverCheck_RequestServerName = LCase(Request.ServerVariables("SERVER_NAME"))  '도메인 이름
		
		'사이트 변수 설정 S
		
		Select Case NaverCheck_RequestServerName

			'Campus 
			Case "campus.megastudy.net", "campus1.megastudy.net", "campus2.megastudy.net", "campus5.megastudy.net", "amsdev.megastudy.net" ,"tcampus.megastudy.net", "mcampus.megastudy.net", "mcampus1.megastudy.net", "mcampus2.megastudy.net", "mcampus5.megastudy.net" , "tmcampus.megastudy.net", "campus3.megastudy.net", "mcampus3.megastudy.net"
	   
				NaverCheck_CampusDetailCode = ""
				detailNaverCheck_CampusDetailCode = "9d6d7d0820f3efbf43bd3fd3ba73ca6f7449e3b5"
			
			'강남 재수관	http://gangnam.megastudy.net	http://mgangnam.megastudy.net
			Case "gangnam.megastudy.net", "mgangnam.megastudy.net", "tgangnam.megastudy.net", "tmgangnam.megastudy.net"

				NaverCheck_CampusDetailCode = "CD0206"
				detailNaverCheck_CampusDetailCode = "40010b1608b9242b424c2eecb1c110aeb3839cc6"
			
			'강남 중.고등관	http://gangnams.megastudy.net	http://mgangnams.megastudy.net
			Case "gangnams.megastudy.net", "mgangnams.megastudy.net", "tgangnams.megastudy.net", "tmgangnams.megastudy.net"
			
				NaverCheck_CampusDetailCode = "CD0205"
				
			'강동 재수관	http://gangdong.megastudy.net	http://mgangdong.megastudy.net
			Case "gangdong.megastudy.net", "mgangdong.megastudy.net", "tgangdong.megastudy.net", "tmgangdong.megastudy.net"
						
				NaverCheck_CampusDetailCode = "CD0240"

			'강동 초.중.고	http://gangdongs.megastudy.net	http://mgangdongs.megastudy.net
			Case "gangdongs.megastudy.net", "mgangdongs.megastudy.net", "tgangdongs.megastudy.net", "tmgangdongs.megastudy.net"
				NaverCheck_CampusDetailCode = "CD0215"
			
			'강북 재수관	http://gangbuk.megastudy.net	http://mgangbuk.megastudy.net
			Case "gangbuk.megastudy.net", "mgangbuk.megastudy.net", "tgangbuk.megastudy.net", "tmgangbuk.megastudy.net"
			
				NaverCheck_CampusDetailCode = "CD0210"
				detailNaverCheck_CampusDetailCode = "cbf633f2c2aecda3a27e88652667839abf825c93"

			'강북 고등관	http://gangbuks.megastudy.net	http://mgangbuks.megastudy.net
			Case "gangbuks.megastudy.net", "mgangbuks.megastudy.net", "tgangbuks.megastudy.net", "tmgangbuks.megastudy.net"
			
				NaverCheck_CampusDetailCode = "CD0209"
				detailNaverCheck_CampusDetailCode = "82c487da1c8ef63a7b0562d1d3d16a75f5753ca1"

			'노량진 재수관	http://noryangjin.megastudy.net	http://mnoryangjin.megastudy.net
			Case "noryangjin.megastudy.net", "mnoryangjin.megastudy.net", "tnoryangjin.megastudy.net", "tmnoryangjin.megastudy.net"
							
				NaverCheck_CampusDetailCode = "CD0211"
				detailNaverCheck_CampusDetailCode = "2d723a4927de85db9188120a8df7a2310dd9a670"

			'노량진 단과	http://nrj.megastudy.net	http://mnrj.megastudy.net
			Case "nrj.megastudy.net", "mnrj.megastudy.net","tnrj.megastudy.net", "tmnrj.megastudy.net"
								
				NaverCheck_CampusDetailCode = "CD0212"

			'서초 재수관	http://seocho.megastudy.net	http://mseocho.megastudy.net
			Case "seocho.megastudy.net", "mseocho.megastudy.net", "tseocho.megastudy.net", "tmseocho.megastudy.net"
							
				NaverCheck_CampusDetailCode = "CD0208"
				detailNaverCheck_CampusDetailCode = "dcf7f0020252811c17ed535a57eedfbe55ef7571"

			'서초 고등관	http://seochos.megastudy.net	http://mseochos.megastudy.net
			Case "seochos.megastudy.net", "mseochos.megastudy.net", "tseochos.megastudy.net", "tmseochos.megastudy.net"
				
				NaverCheck_CampusDetailCode = "CD0207"
				
			'성북 재수관	http://seongbuk.megastudy.net	http://mseongbuk.megastudy.net
			Case "seongbuk.megastudy.net", "mseongbuk.megastudy.net", "tseongbuk.megastudy.net", "tmseongbuk.megastudy.net"

				NaverCheck_CampusDetailCode = "CD0239"
				
			'성북 초.중.고	http://seongbuks.megastudy.net	 http://mseongbuks.megastudy.net
			Case "seongbuks.megastudy.net", "mseongbuks.megastudy.net", "tseongbuks.megastudy.net", "tmseongbuks.megastudy.net"
				
				NaverCheck_CampusDetailCode = "CD0214"
				detailNaverCheck_CampusDetailCode = "ef24c3e46a702277ac0787fa8afcebb03e836737"
			
			'신촌 재수정규반	http://sinchon.megastudy.net	http://msinchon.megastudy.net
			Case "sinchon.megastudy.net", "msinchon.megastudy.net", "tsinchon.megastudy.net", "tmsinchon.megastudy.net"

				NaverCheck_CampusDetailCode = "CD0213"
				detailNaverCheck_CampusDetailCode = "c864f0d1a48a920d36d14e16d45e7ab54c6a555d"

			'평촌 재수관	http://pyeongchon.megastudy.net	http://mpyeongchon.megastudy.net
			Case "pyeongchon.megastudy.net", "mpyeongchon.megastudy.net" , "tpyeongchon.megastudy.net", "tmpyeongchon.megastudy.net"

				NaverCheck_CampusDetailCode = "CD0217"
				detailNaverCheck_CampusDetailCode = "0820ea8324c4515e31b5eb346d086676a5c3af47"

			'평촌 고등관	http://pyeongchons.megastudy.net	http://mpyeongchons.megastudy.net
			Case "pyeongchons.megastudy.net", "mpyeongchons.megastudy.net","tpyeongchons.megastudy.net", "tmpyeongchons.megastudy.net"

				NaverCheck_CampusDetailCode = "CD0216"

			'양지 본원 재수정규반	http://yangji.megastudy.net	http://myangji.megastudy.net
			Case "yangji.megastudy.net", "myangji.megastudy.net", "tyangji.megastudy.net", "tmyangji.megastudy.net"	

				NaverCheck_CampusDetailCode = "CD0220"
				detailNaverCheck_CampusDetailCode = "dbd011ceaa6cb26988c49e81a65fe47919a639ab"

			'양지 신관/서초기숙 재수정규반	http://yangjim.megastudy.net	http://myangjim.megastudy.net
			Case "yangjim.megastudy.net", "myangjim.megastudy.net"	, "tyangjim.megastudy.net", "tmyangjim.megastudy.net", "seochob.megastudy.net", "mseochob.megastudy.net"	, "tseochob.megastudy.net", "tmseochob.megastudy.net"	

				NaverCheck_CampusDetailCode = "CD0243"
				detailNaverCheck_CampusDetailCode = "f426f87682d3a84296ee912a744592765dd9425a"

			'부천	재수정규반	http://bucheon.megastudy.net
			Case "bucheon.megastudy.net", "mbucheon.megastudy.net", "tbucheon.megastudy.net", "tmbucheon.megastudy.net"	

				NaverCheck_CampusDetailCode = "CD0251"
				detailNaverCheck_CampusDetailCode = "0ade76b50f0293213b75e09fadf8884ebfae23c5"
			
			'일산	재수정규반	http://ilsan.megastudy.net
			Case "ilsan.megastudy.net", "milsan.megastudy.net", "tilsan.megastudy.net", "tmilsan.megastudy.net"	

				NaverCheck_CampusDetailCode = "CD0255"
				detailNaverCheck_CampusDetailCode = "7b5099ed5e72ad2a983d07fe13196087d7f7bd4b"

			'분당	재수정규반	http://bundang.megastudy.net
			Case "bundang.megastudy.net", "mbundang.megastudy.net", "tbundang.megastudy.net", "tmbundang.megastudy.net"	

				NaverCheck_CampusDetailCode = "CD0253"
				detailNaverCheck_CampusDetailCode = "be67589cde709fa3569b934ba920a74e6d66edaa"

			'러셀대치 http://campus.magastudy.net/russel/
            Case "russeldc.megastudy.net", "mrusseldc.megastudy.net"

                 NaverCheck_CampusDetailCode = "CD0241"

           '러셀분당 http://campus.magastudy.net/russel_bundang/
           Case "russelbd.megastudy.net", "mrusselbd.megastudy.net"

                NaverCheck_CampusDetailCode = "CD0244"

           '러셀목동 http://campus.magastudy.net/russel_mokdong/
            Case "russelmd.megastudy.net", "mrusselmd.megastudy.net"

                NaverCheck_CampusDetailCode = "CD0245"

			'양주 초.중.고	http://yangju.megastudy.net	 http://myangju.megastudy.net
			Case "yangju.megastudy.net", "myangju.megastudy.net", "tyangju.megastudy.net", "tmyangju.megastudy.net"
				
				NaverCheck_CampusDetailCode = "CD0259"

			'송파 재수관	http://songpa.megastudy.net	http://msongpa.megastudy.net
			Case "songpa.megastudy.net", "msongpa.megastudy.net", "tsongpa.megastudy.net", "tmsongpa.megastudy.net"
						
				NaverCheck_CampusDetailCode = "CD0276"
				detailNaverCheck_CampusDetailCode = "005c26b930b430e9b0af9598a47395b4df02abe2"

			'안성기숙	http://ansung.megastudy.net	http://mansung.megastudy.net
			Case "ansung.megastudy.net", "mansung.megastudy.net", "tansung.megastudy.net", "tmansung.megastudy.net"	
						
				NaverCheck_CampusDetailCode = "CD0280"				

			Case Else

				NaverCheck_CampusDetailCode = "" 
			
		End Select

		'사이트 변수 설정 E

		'강남 메가스터디학원 재수관 CD0206	'강남 메가스터디학원 중고등관 CD0205	'강북메가스터디학원 재수관 CD0210	'강북메가스터디학원 고등관 CD0209	'노량진 메가스터디학원 재수관 CD0211	'신촌 메가스터디학원 CD0213	'평촌 메가스터디학원 CD0217	'평촌 메가스터디학원 고등관 CD0216	'강동 메가스터디학원 재수관 CD0240	'강동 메가스터디학원 초중고관 CD0215 '송파 메가스터디학원 초중고관 CD0276 
		'<meta name="naver-site-verification" content="6f8d9536b461cf0fb667a1a89afac4a92b8ae961"/> 사용
		'그외  <meta name="naver-site-verification" content="8e19e1ad08a3c6b9abee62054f8eedddd706f119"/>
		' 송파는 005c26b930b430e9b0af9598a47395b4df02abe2 코드로 변경 (2020-12-10)

		If NaverCheck_CampusDetailCode = "CD0206" Or NaverCheck_CampusDetailCode = "CD0205" Or NaverCheck_CampusDetailCode = "CD0210" Or NaverCheck_CampusDetailCode = "CD0209" Or NaverCheck_CampusDetailCode = "CD0211" Or NaverCheck_CampusDetailCode = "CD0213" Or NaverCheck_CampusDetailCode = "CD0217" Or NaverCheck_CampusDetailCode = "CD0216" Or NaverCheck_CampusDetailCode = "CD0240" Or NaverCheck_CampusDetailCode = "CD0215" Or NaverCheck_CampusDetailCode = "CD0259" Or NaverCheck_CampusDetailCode = "CD0251" Or NaverCheck_CampusDetailCode = "CD0253" Or NaverCheck_CampusDetailCode = "CD0208" Or NaverCheck_CampusDetailCode = "CD0207"  Or NaverCheck_CampusDetailCode = "CD0239"  Or NaverCheck_CampusDetailCode = "CD0214" Or NaverCheck_CampusDetailCode = "CD0255" Or NaverCheck_CampusDetailCode = "CD0220" Or NaverCheck_CampusDetailCode = "CD0243" Then 
			'isNaverCheck_CampusDetailCode = "1"
		ElseIf NaverCheck_CampusDetailCode = "CD0276" Then
			'isNaverCheck_CampusDetailCode = "2"
		Else 
			'isNaverCheck_CampusDetailCode = "0" 
		End If 

		'GetCampusNaverCheckMeta =  isNaverCheck_CampusDetailCode
		' 학원별로 메타태그 코드값 변경됨 (2021-03-16)
		GetCampusNaverCheckMeta =  detailNaverCheck_CampusDetailCode
	End Function 
	'네이버 메타 체크 가져오기 함수 E




	'네이버 메타 설명 가져오기 함수 가져오기 함수 S 
	Function GetCampusNaverDescMeta()
		
		NaverCheck_CampusDetailDesc = ""

		NaverCheck_RequestServerName = LCase(Request.ServerVariables("SERVER_NAME"))  '도메인 이름
		
		'사이트 변수 설정 S
		
		Select Case NaverCheck_RequestServerName

			'Campus 
			Case "campus.megastudy.net", "campus1.megastudy.net", "campus2.megastudy.net", "campus5.megastudy.net", "amsdev.megastudy.net" ,"tcampus.megastudy.net", "mcampus.megastudy.net", "mcampus1.megastudy.net", "mcampus2.megastudy.net", "mcampus5.megastudy.net" , "tmcampus.megastudy.net"
	   
				NaverCheck_CampusDetailDesc = "대입성공의 모든 공식, 본사 직영, 서울/경기재수학원, 입시/수능전문, 수준별 맞춤합격시스템, 팀플장학혜택"
			
			'강남 재수관	http://gangnam.megastudy.net	http://mgangnam.megastudy.net
			Case "gangnam.megastudy.net", "mgangnam.megastudy.net", "tgangnam.megastudy.net", "tmgangnam.megastudy.net"

				NaverCheck_CampusDetailDesc = "강남/대치/서초 재수학원, 수능/입시전문, 팀플장학혜택, 논술대비"
			
			'강남 중.고등관	http://gangnams.megastudy.net	http://mgangnams.megastudy.net
			Case "gangnams.megastudy.net", "mgangnams.megastudy.net", "tgangnams.megastudy.net", "tmgangnams.megastudy.net"
			
				NaverCheck_CampusDetailDesc = "강남구 대치동 위치, 재학생전문학원, 내신/수능 대비, 중/고등부 연합반, 입시학원"
				
			'강동 재수관	http://gangdong.megastudy.net	http://mgangdong.megastudy.net
			Case "gangdong.megastudy.net", "mgangdong.megastudy.net", "tgangdong.megastudy.net", "tmgangdong.megastudy.net"
						
				NaverCheck_CampusDetailDesc = "강동/송파/잠실 재수학원, 수능/입시전문, 팀플장학혜택, 멘토링 프로그램, 재수정규반"

			'강동 초.중.고	http://gangdongs.megastudy.net	http://mgangdongs.megastudy.net
			Case "gangdongs.megastudy.net", "mgangdongs.megastudy.net", "tgangdongs.megastudy.net", "tmgangdongs.megastudy.net"

				NaverCheck_CampusDetailDesc = "초중고 재학생 전문 학원, 독학재수, 재수예체능, 단과반 운영, 내신/수능 완벽 대비"
			
			'강북 재수관	http://gangbuk.megastudy.net	http://mgangbuk.megastudy.net
			Case "gangbuk.megastudy.net", "mgangbuk.megastudy.net", "tgangbuk.megastudy.net", "tmgangbuk.megastudy.net"
			
				NaverCheck_CampusDetailDesc = "강북 입시의 중심! 강북/노원/상계 재수학원, 수능/입시전문, 팀플장학혜택, 논술대비"

			'강북 고등관	http://gangbuks.megastudy.net	http://mgangbuks.megastudy.net
			Case "gangbuks.megastudy.net", "mgangbuks.megastudy.net", "tgangbuks.megastudy.net", "tmgangbuks.megastudy.net"
			
				NaverCheck_CampusDetailDesc = "고등 전문, 재수예체능, 입시전략, 내신/수능 동시대비, 노원구 상계동 위치"

			'노량진 재수관	http://noryangjin.megastudy.net	http://mnoryangjin.megastudy.net
			Case "noryangjin.megastudy.net", "mnoryangjin.megastudy.net", "tnoryangjin.megastudy.net", "tmnoryangjin.megastudy.net"
							
				NaverCheck_CampusDetailDesc = "메가스터디학원 우수 강사진, 남녀분리 자습관, 검증된 관리, 입시전문 담임"

			'노량진 단과	http://nrj.megastudy.net	http://mnrj.megastudy.net
			Case "nrj.megastudy.net", "mnrj.megastudy.net","tnrj.megastudy.net", "tmnrj.megastudy.net"
								
				NaverCheck_CampusDetailDesc = "동작구 노량진동 위치, 노량진 단과"

			'서초 재수관	http://seocho.megastudy.net	http://mseocho.megastudy.net
			Case "seocho.megastudy.net", "mseocho.megastudy.net", "tseocho.megastudy.net", "tmseocho.megastudy.net"
							
				NaverCheck_CampusDetailDesc = "목표를 완성하는, 완벽한 N수의 시작! 수능전문, 수능전문담임, 팀플장학"

			'서초 고등관	http://seochos.megastudy.net	http://mseochos.megastudy.net
			Case "seochos.megastudy.net", "mseochos.megastudy.net", "tseochos.megastudy.net", "tmseochos.megastudy.net"
				
				NaverCheck_CampusDetailDesc = "수능 전과목 만점자 배출! 의대입시전문, 자연계 최상위권 대입, 팀플장학혜택, 재수정규반"
				
			'성북 재수관	http://seongbuk.megastudy.net	http://mseongbuk.megastudy.net
			Case "seongbuk.megastudy.net", "mseongbuk.megastudy.net", "tseongbuk.megastudy.net", "tmseongbuk.megastudy.net"

				NaverCheck_CampusDetailDesc = "성북/강북/동대문 재수학원, 수능/입시전문, 논술대비, 팀플장학혜택, 퀀텀점프 성적향상, 재수정규반"
				
			'성북 초.중.고	http://seongbuks.megastudy.net	 http://mseongbuks.megastudy.net
			Case "seongbuks.megastudy.net", "mseongbuks.megastudy.net", "tseongbuks.megastudy.net", "tmseongbuks.megastudy.net"
				
				NaverCheck_CampusDetailDesc = "2023 서울대 의대 수시합격. 고등 단과 국수영탐. 고등 학교별 전문내신단과. 고3 내신&수능 ALL KILL 수업. 중등/초등 수학/영어/과학"
			
			'신촌 재수정규반	http://sinchon.megastudy.net	http://msinchon.megastudy.net
			Case "sinchon.megastudy.net", "msinchon.megastudy.net", "tsinchon.megastudy.net", "tmsinchon.megastudy.net"

				NaverCheck_CampusDetailDesc = "검증된 입시 전문 담임 관리제! 홍대/신촌/합정 재수학원, 수능/입시전문, 팀플장학혜택, 논술대비"

			'평촌 재수관	http://pyeongchon.megastudy.net	http://mpyeongchon.megastudy.net
			Case "pyeongchon.megastudy.net", "mpyeongchon.megastudy.net" , "tpyeongchon.megastudy.net", "tmpyeongchon.megastudy.net"

				NaverCheck_CampusDetailDesc = "성적이 오르는, 합격을 위한 바른 관리! 안양/수원/안산/과천 재수학원, 수능/입시전문, 팀플장학혜택, 논술대비"

			'평촌 고등관	http://pyeongchons.megastudy.net	http://mpyeongchons.megastudy.net
			Case "pyeongchons.megastudy.net", "mpyeongchons.megastudy.net","tpyeongchons.megastudy.net", "tmpyeongchons.megastudy.net"

				NaverCheck_CampusDetailDesc = "안양, 평촌  재학생전문학원, 내신/수능 대비, 고등부 연합반, 단과반 운영"

			'양지 기숙 본원 재수정규반	http://yangji.megastudy.net	http://myangji.megastudy.net
			Case "yangji.megastudy.net", "myangji.megastudy.net", "tyangji.megastudy.net", "tmyangji.megastudy.net"	

				NaverCheck_CampusDetailDesc = "국내유일 남녀분반 문이과 상위권전문관, 메가스터디교육 직영, 수능/입시전문, 팀플장학혜택, 논술대비, 용인시 처인구 원삼면 위치"

			' 서초 기숙 양지 신관 재수정규반	http://yangjim.megastudy.net	http://myangjim.megastudy.net
			Case "yangjim.megastudy.net", "myangjim.megastudy.net"	, "tyangjim.megastudy.net", "tmyangjim.megastudy.net" ,"seochob.megastudy.net", "mseochob.megastudy.net"	, "tseochob.megastudy.net", "tmseochob.megastudy.net"		

				NaverCheck_CampusDetailDesc = "(구) 서초 메가스터디 기숙학원, 메가스터디 직영 기숙학원, 메디컬 합격을 위한 최선의 선택, 팀플장학혜택, 오직 자연계 합격을 연구, 자연계열 학생들의 취약점에 Focus를 맞춘 수업, 국어 특화수업/수학, 탐구과목 수준별 수업, 축적된 성적향상 노하우, Miracle 팀플 의대반" 
			
			'부천	재수정규반	http://bucheon.megastudy.net
			Case "bucheon.megastudy.net", "mbucheon.megastudy.net", "tbucheon.megastudy.net", "tmbucheon.megastudy.net"	

				NaverCheck_CampusDetailDesc = "독한공부 확실한 결과! 인천/시흥/부평/김포 재수학원, 수능/입시전문, 팀플장학혜택, 논술대비"
			
			'일산	재수정규반	http://ilsan.megastudy.net
			Case "ilsan.megastudy.net", "milsan.megastudy.net", "tilsan.megastudy.net", "tmilsan.megastudy.net"	

				NaverCheck_CampusDetailDesc = "맞춤관리·성적향상의 중심! 일산/파주/김포 재수학원, 수능/입시전문, 팀플장학혜택, 논술대비"

			'분당	재수정규반	http://bundang.megastudy.net
			Case "bundang.megastudy.net", "mbundang.megastudy.net", "tbundang.megastudy.net", "tmbundang.megastudy.net"	

				NaverCheck_CampusDetailDesc = "수준별 맞춤 학습, 철저한 학생 관리! 용인/수지/수원/성남 재수학원, 수능/입시전문, 팀플장학혜택, 논술대비"

			'러셀대치 http://campus.magastudy.net/russel/
            Case "russeldc.megastudy.net", "mrusseldc.megastudy.net"

                 NaverCheck_CampusDetailDesc = ""

           '러셀분당 http://campus.magastudy.net/russel_bundang/
           Case "russelbd.megastudy.net", "mrusselbd.megastudy.net"

                NaverCheck_CampusDetailDesc = ""

           '러셀목동 http://campus.magastudy.net/russel_mokdong/
            Case "russelmd.megastudy.net", "mrusselmd.megastudy.net"

                NaverCheck_CampusDetailDesc = ""

			'양주 초.중.고	http://yangju.megastudy.net	 http://myangju.megastudy.net
			Case "yangju.megastudy.net", "myangju.megastudy.net", "tyangju.megastudy.net", "tmyangju.megastudy.net"
				
				NaverCheck_CampusDetailDesc = "양주/의정부/동두천/경기지역 초중고 입시전문, 본사직영, 코딩특강, 내신/수능 완벽대비, 독학재수반"
			
			'운정 초.중.고	http://unjung.megastudy.net	http://munjung.megastudy.net/
			Case "unjung.megastudy.net", "munjung.megastudy.net", "tunjung.megastudy.net", "tmunjung.megastudy.net"
				NaverCheck_CampusDetailDesc = "운정/파주/교하지역 초중고 입시전문, 독학재수, 코딩특강, 내신/수능 완벽 대비"

			'송파 재수관	http://songpa.megastudy.net	http://msongpa.megastudy.net
			Case "songpa.megastudy.net", "msongpa.megastudy.net", "tsongpa.megastudy.net", "tmsongpa.megastudy.net"						
				NaverCheck_CampusDetailDesc = "맞춤 관리를 통한 최상의 결과! 강남/송파/강동/광진/구리/하남/남양주 재수학원, 수능/입시전문, 팀플장학혜택, 논술대비"

			'안성기숙 	http://ansung.megastudy.net	http://mansung.megastudy.net
			Case "ansung.megastudy.net", "mansung.megastudy.net", "tansung.megastudy.net", "tmansung.megastudy.net"						
				NaverCheck_CampusDetailDesc = "방법이 옳으면 성적이 오른다! 안성맞춤 성적상승, 서울/경기재수학원, 본사직영 통학/기숙학원, 수능성적 상승 전문, 팀플장학혜택"

			Case Else

				NaverCheck_CampusDetailDesc = "대입성공의 모든 공식, 본사 직영, 서울/경기재수학원, 입시/수능전문, 수준별 맞춤합격시스템, 팀플장학혜택"
			
		End Select

		'사이트 변수 설정 E

		'오메가 모의고사 홍보페이지 처리 / 최민석 / 2025-01-10 17:23:50
		If LCase(request.ServerVariables("PATH_INFO")) = "/campus_common/2025/2025_omega/index.asp" Then
			NaverCheck_CampusDetailDesc = "실력과 실전의 간극을 좁히는 오메가! 맞춤 실전 대비로 대입 성공 완성!"
		'메대프 모의고사 홍보페이지 처리 / 최민석 / 2025-01-15
		ElseIf LCase(request.ServerVariables("PATH_INFO")) = "/campus_common/2025/2025_premium/index.asp" Then
			NaverCheck_CampusDetailDesc = "전 과목 전 문항 모의고사로 완벽한 실전 전략 수립, 그 자체로 프리미엄한 모의고사"
		End If

		GetCampusNaverDescMeta =  NaverCheck_CampusDetailDesc
	End Function 
	'네이버 메타 설명 가져오기 함수 E
%>
<!-- favicon -->	
<link rel="shortcut icon" href="https://img.megastudy.net/campus/library/v2015_mob/asset/img/favicon.ico">
<link rel="apple-touch-icon" href="https://img.megastudy.net/campus/library/v2015_mob/asset/img/iphone.png">
<link rel="apple-touch-icon" sizes="76x76" href="https://img.megastudy.net/campus/library/v2015_mob/asset/img/apple-touch-icon-ipad.png">
<link rel="apple-touch-icon" sizes="120x120" href="https://img.megastudy.net/campus/library/v2015_mob/asset/img/apple-touch-icon-iphone4.png">

<%If LCase(request.ServerVariables("PATH_INFO")) = "/campus_common/2025/2025_omega/index.asp" Then%>
	<!--오메가 모의고사 홍보페이지 처리 / 최민석 / 2025-01-10 17:24:56 -->
	<meta name="author" content="메가스터디학원" />
	<meta name="keywords" content="메가스터디학원, 러셀학원, 재수학원, 입시, 재수기숙학원, 메가스터디러셀, 메가스터디, N수, 대입, 수능, 모의고사, 실전 모의고사" />
	<meta property="og:site_name" content="메가스터디학원">
	<meta property="og:image" content="https://img.megastudy.net/campus/library/v2015/library/campus_common/2025/2025_omega/thumb.jpg">
	<meta property="og:url" content="https://campus.megastudy.net/campus_common/2025/2025_omega/index.asp">
<%ElseIf LCase(request.ServerVariables("PATH_INFO")) = "/campus_common/2025/2025_premium/index.asp" Then%>
	<!--메대프 모의고사 홍보페이지 처리 / 최민석 / 2025-01-15 -->
	<meta name="author" content="메가스터디학원" />
	<meta name="keywords" content="메가스터디학원, 러셀학원, 재수학원, 입시, 재수기숙학원, 메가스터디러셀, 메가스터디, N수, 대입, 수능, 모의고사, 더 프리미엄 모의고사" />
	<meta property="og:site_name" content="메가스터디학원">
	<meta property="og:image" content="https://img.megastudy.net/campus/library/v2015/library/campus_common/2025/2025_premium/thumb.jpg">
	<meta property="og:url" content="https://campus.megastudy.net/campus_common/2025/2025_premium/index.asp">
<%Else%>
	<meta property="og:image" content="https://img.megastudy.net/campus/library/v2015/library/intro/main_logo.gif">
	<meta property="og:url" content="https://<%=Request.ServerVariables("SERVER_NAME")%>">
<%End If%>

<meta name="naver-site-verification" content="<%=GetCampusNaverCheckMeta()%>"/>
<meta name="description" content="<%=GetCampusNaverDescMeta()%>"/>
<meta property="og:type" content="website">
<meta property="og:title" content="<%=gMGC_Campus_Title & IIF(GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_")="CD0214", " 초중고관", "") %>">
<meta property="og:description" content="<%=GetCampusNaverDescMeta()%>">
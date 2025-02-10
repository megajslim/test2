<%
'=======================
' 작업일자 : 2020-10-29
' 작업내용 : 메타태그 관련 함수 
'=======================

	'네이버 메타 체크 가져오기 함수 가져오기 함수 S 
	Function GetCampusNaverMetaCode()
		
		ValCampusNaverMetaCode = ""

		NaverCheck_RequestServerName = LCase(Request.ServerVariables("SERVER_NAME"))  '도메인 이름
		
		'사이트 변수 설정 S
		
		Select Case NaverCheck_RequestServerName

			'Campus 
			Case "campus.megastudy.net", "campus1.megastudy.net", "campus2.megastudy.net", "campus5.megastudy.net", "amsdev.megastudy.net" ,"tcampus.megastudy.net", "mcampus.megastudy.net", "mcampus1.megastudy.net", "mcampus2.megastudy.net", "mcampus5.megastudy.net" , "tmcampus.megastudy.net"
	   
				NaverCheck_CampusDetailCode = ""
				ValCampusNaverMetaCode ="74410de4ca8d9dbe64fdb2586d1b3fee54d39d15"
			
			'강남 재수관	http://gangnam.megastudy.net	http://mgangnam.megastudy.net
			Case "gangnam.megastudy.net", "mgangnam.megastudy.net", "tgangnam.megastudy.net", "tmgangnam.megastudy.net"

				NaverCheck_CampusDetailCode = "CD0206"
				ValCampusNaverMetaCode ="a41615b6cf9b9e1fd52f1201905a697cba9c8f16"
			
			'강남 중.고등관	http://gangnams.megastudy.net	http://mgangnams.megastudy.net
			Case "gangnams.megastudy.net", "mgangnams.megastudy.net", "tgangnams.megastudy.net", "tmgangnams.megastudy.net"
			
				NaverCheck_CampusDetailCode = "CD0205"
				
			'강동 재수관	http://gangdong.megastudy.net	http://mgangdong.megastudy.net
			Case "gangdong.megastudy.net", "mgangdong.megastudy.net", "tgangdong.megastudy.net", "tmgangdong.megastudy.net"
						
				NaverCheck_CampusDetailCode = "CD0240"
				ValCampusNaverMetaCode ="6f8d9536b461cf0fb667a1a89afac4a92b8ae961"

			'강동 초.중.고	http://gangdongs.megastudy.net	http://mgangdongs.megastudy.net
			Case "gangdongs.megastudy.net", "mgangdongs.megastudy.net", "tgangdongs.megastudy.net", "tmgangdongs.megastudy.net"
				NaverCheck_CampusDetailCode = "CD0215"
				ValCampusNaverMetaCode ="6f8d9536b461cf0fb667a1a89afac4a92b8ae961"
			
			'강북 재수관	http://gangbuk.megastudy.net	http://mgangbuk.megastudy.net
			Case "gangbuk.megastudy.net", "mgangbuk.megastudy.net", "tgangbuk.megastudy.net", "tmgangbuk.megastudy.net"
			
				NaverCheck_CampusDetailCode = "CD0210"
				ValCampusNaverMetaCode ="98c88d140eda9591c3ebd0f8d8d7ff700501b5da"

			'강북 고등관	http://gangbuks.megastudy.net	http://mgangbuks.megastudy.net
			Case "gangbuks.megastudy.net", "mgangbuks.megastudy.net", "tgangbuks.megastudy.net", "tmgangbuks.megastudy.net"
			
				NaverCheck_CampusDetailCode = "CD0209"
				ValCampusNaverMetaCode ="5789c4342067c899696829e8090c51a3168b63c7"

			'노량진 재수관	http://noryangjin.megastudy.net	http://mnoryangjin.megastudy.net
			Case "noryangjin.megastudy.net", "mnoryangjin.megastudy.net", "tnoryangjin.megastudy.net", "tmnoryangjin.megastudy.net"
							
				NaverCheck_CampusDetailCode = "CD0211"
				ValCampusNaverMetaCode ="2ab4247c650377e549cf20f8d1565fc9f177ac2e"

			'노량진 단과	http://nrj.megastudy.net	http://mnrj.megastudy.net
			Case "nrj.megastudy.net", "mnrj.megastudy.net","tnrj.megastudy.net", "tmnrj.megastudy.net"
								
				NaverCheck_CampusDetailCode = "CD0212"
				ValCampusNaverMetaCode ="6f8d9536b461cf0fb667a1a89afac4a92b8ae961"

			'서초 재수관	http://seocho.megastudy.net	http://mseocho.megastudy.net
			Case "seocho.megastudy.net", "mseocho.megastudy.net", "tseocho.megastudy.net", "tmseocho.megastudy.net"
							
				NaverCheck_CampusDetailCode = "CD0208"
				ValCampusNaverMetaCode ="17e850024fca26f085bb9f3a8c2c5889debb7cc2"

			'서초 고등관	http://seochos.megastudy.net	http://mseochos.megastudy.net
			Case "seochos.megastudy.net", "mseochos.megastudy.net", "tseochos.megastudy.net", "tmseochos.megastudy.net"
				
				NaverCheck_CampusDetailCode = "CD0207"
				ValCampusNaverMetaCode ="6f8d9536b461cf0fb667a1a89afac4a92b8ae961"
				
			'성북 재수관	http://seongbuk.megastudy.net	http://mseongbuk.megastudy.net
			Case "seongbuk.megastudy.net", "mseongbuk.megastudy.net", "tseongbuk.megastudy.net", "tmseongbuk.megastudy.net"

				NaverCheck_CampusDetailCode = "CD0239"
				ValCampusNaverMetaCode ="6f8d9536b461cf0fb667a1a89afac4a92b8ae961"
				
			'성북 초.중.고	http://seongbuks.megastudy.net	 http://mseongbuks.megastudy.net
			Case "seongbuks.megastudy.net", "mseongbuks.megastudy.net", "tseongbuks.megastudy.net", "tmseongbuks.megastudy.net"
				
				NaverCheck_CampusDetailCode = "CD0214"
				ValCampusNaverMetaCode ="a402334b8adf942a2973187d8f59e2ebd954932e"
			
			'신촌 재수정규반	http://sinchon.megastudy.net	http://msinchon.megastudy.net
			Case "sinchon.megastudy.net", "msinchon.megastudy.net", "tsinchon.megastudy.net", "tmsinchon.megastudy.net"

				NaverCheck_CampusDetailCode = "CD0213"
				ValCampusNaverMetaCode ="d6fb267b194e1b45423f3babcb8b260cb906b38b"

			'평촌 재수관	http://pyeongchon.megastudy.net	http://mpyeongchon.megastudy.net
			Case "pyeongchon.megastudy.net", "mpyeongchon.megastudy.net" , "tpyeongchon.megastudy.net", "tmpyeongchon.megastudy.net"

				NaverCheck_CampusDetailCode = "CD0217"
				ValCampusNaverMetaCode ="8bec87edf0992144f5ccac972f8c97c0505c745f"

			'평촌 고등관	http://pyeongchons.megastudy.net	http://mpyeongchons.megastudy.net
			Case "pyeongchons.megastudy.net", "mpyeongchons.megastudy.net","tpyeongchons.megastudy.net", "tmpyeongchons.megastudy.net"

				NaverCheck_CampusDetailCode = "CD0216"

			'양지 본원 재수정규반	http://yangji.megastudy.net	http://myangji.megastudy.net
			Case "yangji.megastudy.net", "myangji.megastudy.net", "tyangji.megastudy.net", "tmyangji.megastudy.net"	

				NaverCheck_CampusDetailCode = "CD0220"
				ValCampusNaverMetaCode ="c751b91ea40aea2d3c24ab71036849b47f909987"

			'양지 신관/서초기숙 재수정규반	http://yangjim.megastudy.net	http://myangjim.megastudy.net
			Case "yangjim.megastudy.net", "myangjim.megastudy.net"	, "tyangjim.megastudy.net", "tmyangjim.megastudy.net", "seochob.megastudy.net", "mseochob.megastudy.net"	, "tseochob.megastudy.net", "tmseochob.megastudy.net"	

				NaverCheck_CampusDetailCode = "CD0243" 
				NaverCheck_CampusDetailCode = "CD0220"
				ValCampusNaverMetaCode ="76e8e6074fd2a57e73e6da6351c230058112750e"

			'부천	재수정규반	http://bucheon.megastudy.net
			Case "bucheon.megastudy.net", "mbucheon.megastudy.net", "tbucheon.megastudy.net", "tmbucheon.megastudy.net"	

				NaverCheck_CampusDetailCode = "CD0251"
				ValCampusNaverMetaCode ="e7e08b2e2574aa0d59a65bfe8d310a688e9b024a"
			
			'일산	재수정규반	http://ilsan.megastudy.net
			Case "ilsan.megastudy.net", "milsan.megastudy.net", "tilsan.megastudy.net", "tmilsan.megastudy.net"	

				NaverCheck_CampusDetailCode = "CD0255"
				ValCampusNaverMetaCode ="83dbcc4e4f09b3ee355f1b82d889002e94e968e8"

			'분당	재수정규반	http://bundang.megastudy.net
			Case "bundang.megastudy.net", "mbundang.megastudy.net", "tbundang.megastudy.net", "tmbundang.megastudy.net"	

				NaverCheck_CampusDetailCode = "CD0253"
				ValCampusNaverMetaCode ="79102f7dfd4479473dfd456b219446536eea9587"

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
				ValCampusNaverMetaCode = "6f8d9536b461cf0fb667a1a89afac4a92b8ae961"

			'송파 재수관	http://songpa.megastudy.net	http://msongpa.megastudy.net
			Case "songpa.megastudy.net", "msongpa.megastudy.net", "tsongpa.megastudy.net", "tmsongpa.megastudy.net"
						
				NaverCheck_CampusDetailCode = "CD0276"
				ValCampusNaverMetaCode ="0973e8a7fd4fc33abe582c993270d035c00d6258"

			'안성기숙	http://ansung.megastudy.net	http://mansung.megastudy.net
			Case "ansung.megastudy.net", "mansung.megastudy.net", "tansung.megastudy.net", "tmansung.megastudy.net"	
						
				NaverCheck_CampusDetailCode = "CD0280"		

			Case Else

				NaverCheck_CampusDetailCode = "" 
				'ValCampusNaverMetaCode ="8e19e1ad08a3c6b9abee62054f8eedddd706f119"
				ValCampusNaverMetaCode ="74410de4ca8d9dbe64fdb2586d1b3fee54d39d15"
			
		End Select

		'사이트 변수 설정 E

		
		'1	공통	메가스터디학원	서울/경기 재수종합학원, 입시/수능 전문, 2018재수정규반, 의약학전문관, 본사직영 기숙학원, 팀플 장학 혜택	이미 등록되어 있음
		'2	강남	강남 메가스터디학원	재수종합학원, 2018 재수정규반, 수능/입시전문, 논술완벽대비, 팀플 장학 혜택, 강남구 대치동 위치 	<meta name="naver-site-verification" content="6f8d9536b461cf0fb667a1a89afac4a92b8ae961"/>
		'3	강동	강동 메가스터디학원 재수관	재수종합학원, 2018 재수정규반, 수능/입시전문, 논술완벽대비, 팀플 장학 혜택, 강동구 명일동 위치	<meta name="naver-site-verification" content="d33c43ee68f327ae0b169454a2bcd07f04d121c8"/>
		'4	강북	강북 메가스터디학원 재수관	재수종합학원, 2018 재수정규반, 수능/입시전문, 논술완벽대비, 팀플 장학 혜택, 노원구 상계동 위치	<meta name="naver-site-verification" content="d33c43ee68f327ae0b169454a2bcd07f04d121c8"/>
		'5	노량진	노량진 메가스터디학원	재수종합학원, 2018 재수정규반, 수능/입시전문, 논술완벽대비, 팀플 장학 혜택, 동작구 노량진동 위치	<meta name="naver-site-verification" content="d33c43ee68f327ae0b169454a2bcd07f04d121c8"/>
		'6	신촌	신촌 메가스터디학원	재수종합학원, 2018 재수정규반, 입시컨설팅, 전문 입시담임제, 팀플 장학 혜택, 서대문구 창천동 위치	<meta name="naver-site-verification" content="d33c43ee68f327ae0b169454a2bcd07f04d121c8"/>
		'7	서초	서초 메가스터디학원 의약학전문관	의대 입시전문, 자연계 최상위권, 2018 재수정규반, 팀플 장학 혜택, 서초구 서초동 위치	<meta name="naver-site-verification" content="d33c43ee68f327ae0b169454a2bcd07f04d121c8"/>
		'8	성북	성북 메가스터디학원 재수관	재수종합학원, 2018 재수정규반, 수능/입시전문, 논술완벽대비, 팀플 장학 혜택, 성북구 동소문동 위치	<meta name="naver-site-verification" content="d33c43ee68f327ae0b169454a2bcd07f04d121c8"/>
		'9	부천	부천 메가스터디학원	신개념 융합형 재수종합학원, 2018 재수정규반, 팀플 장학 혜택, 부천/인천 지역	<meta name="naver-site-verification" content="d33c43ee68f327ae0b169454a2bcd07f04d121c8"/>
		'10	분당	분당 메가스터디학원	신개념 융합형 재수종합학원, 2018 재수정규반, 팀플 장학 혜택, 성남시 분당구 위치	<meta name="naver-site-verification" content="d33c43ee68f327ae0b169454a2bcd07f04d121c8"/>
		'11	일산	일산 메가스터디학원	일산, 파주 신개념 융합형 재수종합학원, 2018 재수정규반, 팀플반 장학 혜택, 고양시 일산동구 위치	<meta name="naver-site-verification" content="d33c43ee68f327ae0b169454a2bcd07f04d121c8"/>
		'12	평촌	평촌 메가스터디학원	신개념 융합형 재수종합학원, 2018 재수정규반, 팀플 장학 혜택, 안양, 평촌 지역	<meta name="naver-site-verification" content="d33c43ee68f327ae0b169454a2bcd07f04d121c8"/>


		'13	양지기숙	양지 메가스터디 기숙학원 	국내 최대 규모 상위권 전문 기숙학원, 2018 재수정규반, 메가스터디 직영, 용인시 처인구 백암면 위치	<meta name="naver-site-verification" content="f270d91e9bb1d952f4b102b94c2ef41f60d5ba30"/>
		'14	서초기숙	서초 메가스터디 기숙학원	자연계열 프리미엄 재수기숙학원, 최상위권 의약학전문관, 2018 재수정규반, 용인시 처인구 양지면 위치	<meta name="naver-site-verification" content="f270d91e9bb1d952f4b102b94c2ef41f60d5ba30"/>
		'15	강북재학	강북메가스터디학원 고등관	중고등 재학생전문, 재수예체능관 입시전략, 내신/수능 동시대비, 노원구 상계동 위치	<meta name="naver-site-verification" content="f270d91e9bb1d952f4b102b94c2ef41f60d5ba30"/>
		'16	강동재학	강동 메가스터디학원 고등관	재학생전문학원, 초중고 재학생반, 단과반 운영, 내신/수능 대비, 강동구 길동 위치	<meta name="naver-site-verification" content="f270d91e9bb1d952f4b102b94c2ef41f60d5ba30"/>
		'17	성북재학	성북 메가스터디학원 고등관	초중고 재학생/재수예체능 전문, 내신/수능 완벽대비, 성신여대입구역 5번출구 위치	<meta name="naver-site-verification" content="f270d91e9bb1d952f4b102b94c2ef41f60d5ba30"/>

		

		GetCampusNaverMetaCode =  ValCampusNaverMetaCode 
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
				
				NaverCheck_CampusDetailDesc = "수능 전과목 만점자 배출! 의약학전문관, 자연계 최상위권 대입, 팀플장학혜택, 재수정규반"
				
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
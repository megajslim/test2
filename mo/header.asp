<%
'####################################################
' 쿠키값이 사라진 경우 다시세팅하기 시작
'####################################################
If ValueCheck(GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_")) = False Then
	sPathInfo = Request.ServerVariables("PATH_INFO")
	If sPathInfo <> "/index.asp" Then '인트로페이지에서는 체크하지말것
		sSql = ""
		sSql = sSql & vbCrlf & "SELECT STUFF((SELECT ',' + CODE"
		sSql = sSql & vbCrlf & "              FROM   MG_CODE_SUB"
		sSql = sSql & vbCrlf & "              WHERE  KIND = 'campus_detail'"
		sSql = sSql & vbCrlf & "              AND    NAME3 IN ('jaehak', 'nsu', 'common', 'danka', 'gisuk')"
		sSql = sSql & vbCrlf & "              AND    ACTIVE = 'Y'"
		sSql = sSql & vbCrlf & "              FOR    XML PATH('')), 1, 1, '')"
		sCampusCodes = objDb.sqlResult(sSql)
		'**** 공통코드에 등록된 학원코드를 토대로 하여, 현재페이지의 url과 매칭하면서 코드를 찾는다.
		sCurCampusCode = GetCampusCodeByDir(Split(sCampusCodes, ",")) 
		'**** 찾은 코드가 있다면, 해당 학원의 쿠키를 다시 세팅하기 위하여, index.asp 파일로 refer와 함께 리디렉션 시킨다
		If ValueCheck(sCurCampusCode) = True Then 
		
			sRedUrl = "/" & GetCampusDirDetail(sCurCampusCode) & "/index.asp"
			If ValueCheck(Request.QueryString) = True Then
				sPathInfo = sPathInfo & "?" & Request.QueryString
			End If

			sRedUrl = sRedUrl & "?refer=" & Server.URLEncode(sPathInfo)
			Call js_redirect(sRedUrl, "")
			'println "이렇게 호출함 = " & sRedUrl
			Response.End
		End If
	End If
End If
'####################################################
' 쿠키값이 사라진 경우 다시세팅하기 끝
'####################################################

If GetCampusVarsSiteMGC("_CAMPUS_THEME_") = "red" Then
	Response.Cookies("_CAMPUS_THEME1_") = "first"
	Response.Cookies("_CAMPUS_THEME2_") = "last1"
End If

If GetCampusVarsSiteMGC("_CAMPUS_THEME_") = "blue" Then
	Response.Cookies("_CAMPUS_THEME1_") = "second"
	Response.Cookies("_CAMPUS_THEME2_") = "last2"
End If

If GetCampusVarsSiteMGC("_CAMPUS_THEME_") = "beige" Then
	Response.Cookies("_CAMPUS_THEME1_") = "third"
	Response.Cookies("_CAMPUS_THEME2_") = "last3"
End If

If GetCampusVarsSiteMGC("_CAMPUS_THEME_") = "purple" Then
	Response.Cookies("_CAMPUS_THEME1_") = "four"
	Response.Cookies("_CAMPUS_THEME2_") = "last4"
End If

phonenum="" : phonenum1="" : parentYN = "N" : tmpTopBannerView = "Y"

'CD0220(양지기숙), CD0243(서초기숙), CD0206(강남), CD0253(분당)만 보이게 2021-06-23, 2021-07-05 분당은 숨김
'2024-04-17 CD0206(강남) 부모님 편지 비노출처리 / 최민석
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0206" Then
	phonenum1="02-568-3800"
	phonenum="025683800"
	'parentYN = "Y"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0215" Then
	phonenum1="02-428-8181"
	phonenum="024288181"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0210" Then
	phonenum1="02-935-3378"
	phonenum="029353378"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0211" Then
	phonenum1="02-826-1555"
	phonenum="028261555"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0208" Then
	phonenum1="02-535-5678"
	phonenum="025355678"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0213" Then
	phonenum1="02-3144-1010"
	phonenum="0231441010"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0217" Then
	phonenum1="031-388-3141"
	phonenum="0313883141"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0205" Then
	phonenum1="02-568-3801"
	phonenum="025683801"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0240" Then '강동재수
	phonenum1="02-428-8181"
	phonenum="024288181"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0215" Then '강동재학
	phonenum1="02-3426-8181"
	phonenum="0234268181"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0209" Then
	phonenum1="02-2162-5250"
	phonenum="0221625250"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0212" Then
	phonenum1="02-826-1800"
	phonenum="028261800"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0207" Then
	phonenum1="02-535-8495"
	phonenum="025358495"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0239" Then '성북재수
	phonenum1="02-6264-8001"
	phonenum="0262648001"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0214" Then '성북재학
	phonenum1="02-927-8001"
	phonenum="029278001"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0216" Then
	phonenum1="031-388-3141"
	phonenum="0313883141"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0219" Then
	phonenum1="031-592-3333"
	phonenum="0315923333"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0220" Then	'양지기숙
	phonenum1="031-326-5000"
	phonenum="0313265000"
	parentYN = "Y"
	'tmpTopBannerView = "N"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0243" Then	'서초기숙
	phonenum1="031-797-9332"
	phonenum="0317979332"
	parentYN = "Y"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0280" Then	'안성기숙
	phonenum1="031-8004-1010"
	phonenum="03180041010"
	parentYN = "Y"
	'tmpTopBannerView = "N"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0251" Then '부천
	phonenum1="032-326-9900"
	phonenum="0323269900"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0253" Then '분당
	phonenum1="031-781-3100"
	phonenum="0317813100"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0255" Then '일산
	phonenum1="031-8073-9600"
	phonenum="03180739600"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0259" Then '양주
	phonenum1="031-928-4848"
	phonenum="0319284848"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0261" Then '운정
	phonenum1="031-947-1050"
	phonenum="0310000000"
End If
If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0276" Then '송파재수(상세내용 수정해야함)
	phonenum1="02-428-8181"
	phonenum="024288181"
End If
%>
<!-- 여기 추가해 공통으로 사용하면 되는데 각 학원별 재정,재학 메인페이지에서도 include 되어 있음-_-; 추후 main.asp 에서는 주석처리 필요. 2018-05-04 KMS -->
<!--#Include Virtual = "/common/inc/bindapply_main.asp" --> 
<% 
	If GetCampusVarsSiteMGC("_CAMPUS_") = "intro" Or GetCampusVarsSiteMGC("_CAMPUS_") = "" Then 	'--인트로 
%>
		<!-- Begin page content : intro -->
		<div class="container pad_no" id="top"><!-- style="margin-bottom: 210px;" -->
			<div class="page-header intro-header">
				
				<h1 class="logo"><a href="/index.asp?clear=Y"><span><img src="<%=Application("img_path_mob")%>/common/img/logo.png" alt="메가스터디학원" class="logo-main"></span></a></h1>
				<div class="dropdown pull-right">
				<button id="dLabel" type="button" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
						<img src="<%=Application("img_path_mob")%>/common/img/top_dropDown.png" alt="학원 바로가기" class="top_dropDown">
					</button>

					<table class="table dropdown-menu" role="menu" aria-labelledby="dLabel">
					<tr>
						<th colspan="4" class="nsu">재수</th>
					</tr>					
					<tr>
						<th colspan="4" class="sub-th">통학학원</th>
					</tr>
					<tr>
						<td colspan="2"><a href="/<%=GetCampusDirDetail("CD0206")%>/">강남 팀플전문관</a></td>
						<td colspan="2"><a href="/<%=GetCampusDirDetail("CD0208")%>/">서초 의약학전문관</a></td>
					</tr>
					<tr>
						<td><a href="/<%=GetCampusDirDetail("CD0210")%>/">강북</a></td>
						<td><a href="/<%=GetCampusDirDetail("CD0211")%>/">노량진</a></td>
						<!--<td><a href="/<%=GetCampusDirDetail("CD0239")%>/">성북</a></td>-->
						<td><a href="/<%=GetCampusDirDetail("CD0213")%>/">신촌</a></td>
						<td><a href="/<%=GetCampusDirDetail("CD0276")%>/">송파</a></td>
					</tr>
					<tr>                
						<td><a href="/<%=GetCampusDirDetail("CD0251")%>/">부천 </a></td>
						<td><a href="/<%=GetCampusDirDetail("CD0253")%>/">분당 </a></td>
						<td><a href="/<%=GetCampusDirDetail("CD0255")%>/">일산</a></td>
						<td><a href="/<%=GetCampusDirDetail("CD0217")%>/">평촌</a></td>
					</tr>
					<tr>						
						<td><a href="http://suwonmega.com/" target="_blank">수원<br><span style="font-size:11px">(협력)</span></a></td>
						<td colspan="3"></td>
					</tr>
					<tr>
						<th colspan="4" class="sub-th">기숙학원</th>
					</tr>
					<tr>
						<td colspan="4"><a href="/<%=GetCampusDirDetail("CD0220")%>/">양지 기숙학원<br>(국내최대 상위권전문)</a></td>
					</tr>
					<tr>
						<td colspan="4"><a href="/<%=GetCampusDirDetail("CD0243")%>/">서초 기숙학원<br>(자연계 상위권전문)</a></td>
					</tr>
					<tr>
						<td colspan="4"><a href="/<%=GetCampusDirDetail("CD0280")%>/">안성 기숙학원<br>(상위권전문)</a></td>
					</tr>
					<tr>
						<th colspan="4" class="second jaehak">초중고</th>
					</tr>
					<tr>
						<!--<td><a href="/<%=GetCampusDirDetail("CD0215")%>/">강동</a></td>-->
						<td><a href="/<%=GetCampusDirDetail("CD0209")%>/">강북</a></td>
						<td><a href="/<%=GetCampusDirDetail("CD0214")%>/">성북</a></td>
						<td><a href="/<%=GetCampusDirDetail("CD0259")%>/">양주 <!--<strong class="icon_new">N</strong>--></a></td>
						<!--<td><a href="/<%=GetCampusDirDetail("CD0261")%>/">운정</a></td>-->
						<td></td>
					</tr>
					<tr>
						<th colspan="4" class="consulting">컨설팅</th>
					</tr>
					<tr>
						<td colspan="4"><a href="http://mmcc.megastudy.net" target="_blank">대입컨설팅센터</a></td>
					</tr>
				</table>
				</div>
			</div>
			<div class="clearfix"></div>
			<div class="mask hide"></div>
<% 
	Else	'--학원별 사이트.
		If  HasValue(GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_"), "CD0214,CD0243,CD0220", ",", True) = True Then
			teacher_move_url = "/teacher/teacher.asp"
		Else
			teacher_move_url = "/teacher/teacherN.asp"
		End If
%>

		<!-- Begin page content : 각 학원페이지 -->
		<!--2024-12-26 최민석 메인에는 main_wrap, 그 외 sub_wrap--> 
		<%If InStr(LCase(Request.ServerVariables("PATH_INFO")), "main.asp") > 0 Or InStr(LCase(Request.ServerVariables("PATH_INFO")), "main_n.asp") > 0 Then%>
		<div id="main_wrap">
		<%Else%>
		<div id="sub_wrap">
		<%End If%>
			<!-- 24.10.04(won)	탑배너 슬라이드 적용 > 작업시 div 형태와 순서 유의하여 반영 -->
			<script type="text/javascript" src="/asset/js/slick.js"></script>
			<link rel="stylesheet" type="text/css" href="/asset/css/slick.css">
			<style>
				#slick_top_banner {width:100%; overflow:hidden;}
				#slick_top_banner * {padding:0; margin:0; box-sizing:border-box;}
				#slick_top_banner .slick-slide {border:none; font-size:0;}
				#slick_top_banner .slick-slide img {width:100%; height:auto;}
			</style>
			<%
				if (GetCampusVarsSiteMGC("_CAMPUS_KIND_") = "nsu")  then 
			%>
				<div class="" id="slick_top_banner">
					<%
						if (GetCampusVarsSiteMGC("_CAMPUS_KIND_") = "nsu") AND HasValue(GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_"), "CD0208,CD0210,CD0211,CD0213,CD0276,CD0251,CD0253,CD0217,CD0209,CD0255", ",", True) = True then 
							if tmpTopBannerView = "Y" AND (InStr(LCase(Request.ServerVariables("PATH_INFO")), "main.asp") > 0) then
							'  2026 N수 우선반 통합 설명회 탑배너 
								If (CDbl("2025"&"0103"&"163000") <= CDbl(GetNowDate()) AND CDbl(GetNowDate()) < CDbl("2025"&"0202"&"140000")) Then
					%>
						<!-- 인트로 2026 N수 성공전략 통합 설명회 // -->
						<div class="slickCnt top_bn_area" style="background: #010815;">
							<a href="/event/2026_concert/index.asp">
								<img src="https://img.megastudy.net/campus/library/v2015_mob/asset/img/ban_2026_concert_250103_m.jpg" alt="">
							</a>
						</div>
						<!-- // 인트로 2026 N수 성공전략 통합 설명회 -->
					<%
								End If
							End If
						End If
					%>		


					<%
						if (GetCampusVarsSiteMGC("_CAMPUS_KIND_") = "nsu") AND HasValue(GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_"), "CD0206,CD0208,CD0210,CD0211,CD0213,CD0276,CD0251,CD0253,CD0217,CD0209,CD0255", ",", True) = True then 
							if tmpTopBannerView = "Y" AND (InStr(LCase(Request.ServerVariables("PATH_INFO")), "main.asp") > 0 Or InStr(LCase(Request.ServerVariables("PATH_INFO")), "main_n.asp") > 0) then
							'  2026 N수 우선반 통합 설명회 탑배너 
								If (CDbl("2024"&"1218"&"090000") <= CDbl(GetNowDate()) AND CDbl(GetNowDate()) < CDbl("2024"&"1228"&"140000")) Then
					%>	
								<div class="slickCnt"><a href="/event/2026_early/index.asp" class="dpCheck"><img src="https://img.megastudy.net/campus/library/v2015_mob/asset/img/ban_2026_early_241217_m.jpg" style="width:100%;" class="img-responsive"></a></div>
					<%
								End If
							End If
						End If
					%>						
						

					<!--맞춤학습시스템 / cms / 20241114-->
					<%
						if HasValue(GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_"), "CD0208,CD0210,CD0211,CD0213,CD0276,CD0251,CD0253,CD0217,CD0209,CD0255", ",", True) = True  AND (InStr(LCase(Request.ServerVariables("PATH_INFO")), "main.asp") > 0 Or InStr(LCase(Request.ServerVariables("PATH_INFO")), "main_n.asp") > 0) Then '성북 재학, 기숙 제외
							If (CDbl("2024"&"1228"&"140000") <= CDbl(GetNowDate())) Then' AND CDbl(GetNowDate()) < CDbl("2024"&"1114"&"180000") ) Then
							%>
								<div class="slickCnt top_bn_area" style="background:#0D3161;">
									<a href="/campus_common/2024/2024_custom_pass/index.asp">
										<img src="https://img.megastudy.net/campus/library/v2015_mob/asset/img/ban_241114_ac_m.jpg" alt="맞춤학습시스템" style="width:100%;" class="img-responsive">
									</a>
								</div>

								<!-- #include virtual="/campus_common/intro/index_custom_layerpop.asp" -->
							<%
							End If
						End If
					%>
					<!--//맞춤학습시스템 / cms / 20241114-->
				</div>
			<%
				end if 
			%>
			<script>
				$(function() {
					if($(".slickCnt").length < 1){
						$("#slick_top_banner").hide();
					}else{
						$("#slick_top_banner").slick({
							infinite:true,
							speed:300,
							autoplay:true,
							autoplaySpeed:3000,
							vertical:true,
							adaptiveHeight:false,
							arrows: false, 
							slidesToShow:1,
							slidesToScroll:1,
							dots:false,
						});
					}
					
				});

				function setIntegerImageHeight() {
					$('#slick_top_banner .slickCnt img').each(function () {
						const img = $(this);

						const containerWidth = img.closest('.slickCnt').width();
						const naturalHeight = img[0].naturalHeight;
						const naturalWidth = img[0].naturalWidth;
						const newHeight = Math.round((containerWidth * naturalHeight) / naturalWidth);

						img.css('height', newHeight + 'px');
					});
				}

				setIntegerImageHeight();

				$(window).on('resize', function () {
					setIntegerImageHeight();
				});
			</script>
			<!-- 헤더 영역 -->
			<header id="header">
				<!-- 학원 로고영역 -->
				<div class="top_wrap">
					<a href="/<%=GetCampusDirDetail(GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_"))%>" target="_top">
						<% If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0243" Then '서초기숙 %>
							<img src="<%=Application("img_path_mob")%>/asset/img/main_renew/logo/logo_russel_ns.svg" alt="<%=GetCampusVarsSiteMGC("_CAMPUS_NAME_")%>">
						<% ElseIf GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0220" Then '양지기숙 %>
							<img src="<%=Application("img_path_mob")%>/asset/img/main_renew/logo/logo_russel_hs.svg" alt="<%=GetCampusVarsSiteMGC("_CAMPUS_NAME_")%>">
						<% Else %>
							<img src="<%=Application("img_path_mob")%>/asset/img/main_renew/logo/logo_<%=GetCampusVarsSiteMGC("_CAMPUS_ENG_NAME_")%>.svg" alt="<%=GetCampusVarsSiteMGC("_CAMPUS_NAME_") %>">
						<%End IF%>
					</a>
					<a class="btn_menu" href="javascript:void(0);">메뉴보기</a>
				</div>
				<!-- //학원 로고영역 -->

				<div class="gnb_menu_wrap">
					<div class="overlay"></div>
					
					<div class="gnb_menu ct_inner">
						<div class="gnb_top">
							<div class="gnb_tit">MENU</div> 
							<a class="gnb_btn close">메뉴닫기</a>
						</div>

						<ul class="gnb_menu_list">
							<% If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") = "CD0239" Then '성북 사이트제한 작업(2019.11.15) %>
								<li>
									<a href="/notice/notice_list.asp">공지사항</a>
								</li>
							<% Else %>
								<li>
									<a href="/<%=GetCampusDirDetail(GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_"))%>/greetings.asp">학원소개</a>
								</li>
								<li>
									<a href="<%=sRecruitHref%>">모집안내</a>
								</li>
								<%If True Then' If(GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_") <> "CD0255") Then ' 일산재종 출강강사진 페이지 비노출 / 2024-12-05 / 정민성 |||| 재노출 / 최민석 / 2025-02-06%>
									<li>
										<a href="<%=teacher_move_url%>">선생님</a>
									</li>
								<% End If %>
									<li>
										<a href="/notice/notice_list.asp">공지사항</a>
									</li>
									<li>
										<a href="/campus/photo_life.asp">학원생활</a>
									</li>
								<% If parentYN = "Y" Then '4곳만 노출 %>
									<li>
										<a href="/story/parent.asp">부모님편지</a>
									</li>
								<% End If%>
									<li>
										<a href="/help/qna.asp">온라인 상담</a>
									</li>
							<% End If %>
						</ul>
					</div>
						
					<div class="contact_area">
						<p class="txt">문의/상담 <a class="tel" href="tel:<%=phonenum%>"><%=phonenum1%></a></p>
					</div>
				</div>
			</header>
			<div class="container pad_no" id="top">
			<!-- //탑배너 슬라이드 적용 -->

			<div class="clearfix"></div>
			<div class="mask hide"></div>
<% 
	End If 
%>
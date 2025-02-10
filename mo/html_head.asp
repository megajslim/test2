<!-- #include virtual = "/models/common/meta.asp" -->
<%
'=======================================================================
'업 무 명 : Mobile Campus 상단 기본 페이지
'모듈기능 :
'파 일 명 : html_head.asp
'작성일자 : 2015/10/
'작 성 자 : 
'-----------------------------------------------------------------------
'변경일자   변경자  변동내역
'=======================================================================
'
'=======================================================================
'Response.Expires=-1
	Dim strHtmlTitle
	If Request.ServerVariables("LOCAL_ADDR") = "10.1.3.10" Then
		strHtmlTitle = "DEV-"
	Else
		strHtmlTitle = ""
	End If
%>
<!doctype html>
<html lang="ko">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
	<meta http-equiv="Pragma" content="no-cache" />
	<meta http-equiv="Cache-Control" content="no-cache" />
	<meta http-equiv="Expires" content="-1" />
    <meta name="viewport" content="initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0, user-scalable=no" /> 
    <meta http-equiv=imagetoolbar content=no>	
	<title><%=strHtmlTitle%><%=gMGC_Campus_Title%></title>
	<!-- favicon -->	
	<link rel="shortcut icon" href="https://img.megastudy.net/campus/library/v2015_mob/asset/img/favicon.ico">
	<link rel="apple-touch-icon" href="https://img.megastudy.net/campus/library/v2015_mob/asset/img/iphone.png">
	<link rel="apple-touch-icon" sizes="76x76" href="https://img.megastudy.net/campus/library/v2015_mob/asset/img/apple-touch-icon-ipad.png">
	<link rel="apple-touch-icon" sizes="120x120" href="https://img.megastudy.net/campus/library/v2015_mob/asset/img/apple-touch-icon-iphone4.png">
<%
	If GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_")="CD0239" Then	' 성북재정일 경우 예외처리
%>
	<meta name="robots" content="noindex">
<%
	Else
%>
	<meta name="naver-site-verification" content="<%=GetCampusNaverMetaCode()%>"/>
	<meta name="description" content="<%=GetCampusNaverDescMeta()%>"/>
	<meta property="og:type" content="website">
	<meta property="og:title" content="<%=gMGC_Campus_Title & IIF(GetCampusVarsSiteMGC("_CAMPUS_DETAIL_CODE_")="CD0214", " 초중고관", "")%>">
	<meta property="og:description" content="<%=GetCampusNaverDescMeta()%>">
	<meta property="og:image" content="https://img.megastudy.net/campus/library/v2015/library/intro/main_logo.gif">
	<meta property="og:url" content="https://<%=Request.ServerVariables("SERVER_NAME")%>">
	<!--메타E-->
<%
	End If
%>
    <!-- Bootstrap core CSS -->
    <link href="/common/css/bootstrap.min.css" rel="stylesheet">

    <!-- Custom styles for this template -->
    <link href="/common/css/sticky-footer.css" rel="stylesheet">
    <link href="/common/css/font-awesome/css/font-awesome.min.css" rel="stylesheet">
    <link href="/common/css/style_new.css?v20170925001" rel="stylesheet"><!-- css 새롭게 반영 -->

	<!-- Owl Carousel Assets -->
    <link href="/common/owl-carousel/owl.carousel.css" rel="stylesheet">
    <link href="/common/owl-carousel/owl.theme.css" rel="stylesheet">

	<!-- 메인 리뉴얼 css, js 파일 추가 / 최민석 / 2024-12-20-->
	<!-- main css -->
	<link rel="stylesheet" type="text/css" href="/common/css/main_renew.css">    
	<!-- swiper -->
	<link rel="stylesheet" type="text/css" href="/asset/css/swiper-bundle.min.css">
	<script type="text/javascript" src="/asset/js/swiper-bundle.min.js"></script>
	
	<!--<link rel="stylesheet" href="/common/css/jquery.bxslider.css">-->

    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/html5shiv/3.7.2/html5shiv.min.js"></script>
      <script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
    <![endif]-->


	<!-- Google Tag Manager 2022.05.09 추가 KMS by 웹서비스팀 -->
	<% 
		Dim gtm_site_url_head_val
		IF LCase(Request.ServerVariables("SERVER_NAME")) = "myangji.megastudy.net" then 
			gtm_site_url_head_val = "GTM-P3ZHV4R"
		ElseIf LCase(Request.ServerVariables("SERVER_NAME")) = "mseochob.megastudy.net" then 
			gtm_site_url_head_val = "GTM-P9Z99W2"		'--PC/MO 공통으로 변경. 2024-12-30 KMS
		ElseIf LCase(Request.ServerVariables("SERVER_NAME")) = "mgangnam.megastudy.net" then 
			gtm_site_url_head_val = "GTM-KRCX5JX"
		ElseIf LCase(Request.ServerVariables("SERVER_NAME")) = "mseocho.megastudy.net" then 
			gtm_site_url_head_val = "GTM-WXLSGJD"
		ElseIf LCase(Request.ServerVariables("SERVER_NAME")) = "mgangbuk.megastudy.net" then 
			gtm_site_url_head_val = "GTM-PJMF4SJ"
		ElseIf LCase(Request.ServerVariables("SERVER_NAME")) = "mnoryangjin.megastudy.net" then 
			gtm_site_url_head_val = "GTM-5XH8M3S"
		ElseIf LCase(Request.ServerVariables("SERVER_NAME")) = "msinchon.megastudy.net" then 
			gtm_site_url_head_val = "GTM-NCRX9HX"
		ElseIf LCase(Request.ServerVariables("SERVER_NAME")) = "msongpa.megastudy.net" then 
			gtm_site_url_head_val = "GTM-T38XT82"
		ElseIf LCase(Request.ServerVariables("SERVER_NAME")) = "mbucheon.megastudy.net" then 
			gtm_site_url_head_val = "GTM-TG9GD6Z"
		ElseIf LCase(Request.ServerVariables("SERVER_NAME")) = "mbundang.megastudy.net" then 
			gtm_site_url_head_val = "GTM-NSPJR8M"
		ElseIf LCase(Request.ServerVariables("SERVER_NAME")) = "milsan.megastudy.net" then 
			gtm_site_url_head_val = "GTM-MJLBC49"
		ElseIf LCase(Request.ServerVariables("SERVER_NAME")) = "mpyeongchon.megastudy.net" then 
			gtm_site_url_head_val = "GTM-5CPWDZS"
		ElseIf LCase(Request.ServerVariables("SERVER_NAME")) = "mseongbuks.megastudy.net" then 
			gtm_site_url_head_val = "GTM-K56V8QL"
		ElseIf LCase(Request.ServerVariables("SERVER_NAME")) = "mansung.megastudy.net" then '안성기숙 추가 2022-11-15 sook
			gtm_site_url_head_val = "GTM-KRF5XFB"
		Else
			gtm_site_url_head_val = "GTM-WBGM58T"
		End if
	%>
	<script>(function(w,d,s,l,i){w[l]=w[l]||[];w[l].push({'gtm.start':
	new Date().getTime(),event:'gtm.js'});var f=d.getElementsByTagName(s)[0],
	j=d.createElement(s),dl=l!='dataLayer'?'&l='+l:'';j.async=true;j.src=
	'https://www.googletagmanager.com/gtm.js?id='+i+dl;f.parentNode.insertBefore(j,f);
	})(window,document,'script','dataLayer','<%=gtm_site_url_head_val%>');</script>
	<!-- End Google Tag Manager 2022.05.09 추가 KMS by 웹서비스팀 -->

	<script type="text/javascript" src="/common/js/util.js?v20180605001"></script>
	<script type="text/javascript" src="/common/js/FN_EchoScriptV2.js"></script>
	<!-- #include virtual = "/include/jquery.asp" -->
	<!--<script type="text/javascript" src="/common/js/jquery.bxslider.min.js"></script>-->

	<!-- 모집안내 관련 버튼 -->
	<script type="text/javascript" src="/apply/common/js/bindapply.js"></script>
	<script type="text/javascript">
		function alert_msg(output_msg)
		{
				title_msg = '메가스터디학원';

			if (!output_msg)
				output_msg = '준비중입니다.';

			$("<div></div>").html(output_msg).dialog({
				title: title_msg,
				resizable: false,
				modal: true,
				buttons: {
					"Ok": function()
					{
						$( this ).dialog( "close" );
					}
				}
			});
		}

	//통학학원 입학식  2019-02-13
	$(document).ready(function() {
		if(window.location.pathname.indexOf("/nsu/") >= 0 || window.location.pathname.indexOf("/sinchon/") >= 0) //재정학원만
		{
			var vCurrentDate = new Date();
			var vCheckDateTime = new Date(2019, 1, 28, 17, 00);

			if( vCurrentDate < vCheckDateTime) {
				$(".sub-accept").before("<div><a href='/notice/notice_view.asp?Code=14235'><img src='https://img.megastudy.net/campus/library/v2015_mob/common/img/intro/intro_190218.jpg' alt='통학학원입학식 다시보기' class='img-responsive'></a></div><br />");
			}

			/* 2019-02-20
			var vCheckDateTime1 = new Date(2019, 1, 19, 18, 50);
			var vCheckDateTime2 = new Date(2019, 1, 19, 21, 00);

			if( vCurrentDate < vCheckDateTime1) {
				$(".sub-accept").before("<div><a href='/notice/notice_view.asp?Code=14213'><img src='https://img.megastudy.net/campus/library/v2015_mob/common/img/intro/intro_190211.jpg' alt='통학학원입학식 안내' class='img-responsive'></a></div><br />");
			} else if (vCurrentDate >= vCheckDateTime1 && vCurrentDate < vCheckDateTime2) {
				$(".sub-accept").before("<div><a href='/notice/notice_view.asp?Code=14214'><img src='https://img.megastudy.net/campus/library/v2015_mob/common/img/intro/intro_190211_on.jpg' alt='통학학원입학식 생중계' class='img-responsive'></a></div><br />");
			} */
		}

		/* gnb */
		let menuOpen = false;

		function updateMenuTop() {
			if (menuOpen) {
				const menuTop = $('#header').offset().top + $('#header .top_wrap').outerHeight();
				$('.gnb_menu_wrap').css('top', menuTop);
			}
		}

		$('.top_wrap .btn_menu').click(function () {
			const menuTop = $('#header').offset().top + $('#header .top_wrap').outerHeight();

			$('.gnb_menu_wrap').css({ 'top': menuTop, 'left': '0' });
			$('body').css('overflow', 'hidden');
			menuOpen = true;
		});

		$('.gnb_btn.close').click(function () {
			$('.gnb_menu_wrap').css('left', '100%');
			$('body').css('overflow', 'auto');
			menuOpen = false;
		});

		$(window).resize(function(){
			updateMenuTop();
		});


		/* 푸터 네비게이션 */
		$('.mega_navi .navi_btn').on('click', function () {
			const type = $(this).data('type');

			$('body').css({'overflow' : 'hidden'});

			if (!$(this).hasClass('intro')) {
				$('.navi_layer').slideDown();

				$('.navi_layer .navi_btn').removeClass('active');
				$(`.navi_layer .navi_btn[data-type="${type}"]`).addClass('active');

				$('.navi_menu').hide();
				$(`.navi_menu.${type}`).show();
			}
		});
	});

	</script>
</head>
<body>
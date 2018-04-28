
<!Doctype html>
<!-- #include file="../style/AVEImgClass.asp"-->
<!-- #include file="../search/connect.asp"-->
<!-- #include file="../site/lang.asp" -->
<!-- #include file="../style/Functions.asp" -->
<!-- #include file="../style/class.json.asp" -->
<!-- #include file="../style/class.json.util.asp" -->
<html>
<head>
<title>Excel to JSON Demo</title>

<link rel="stylesheet" type="text/css" href="../Style/plugin/animation/css/animate.min.css">

<%
  call initJquerySednaicons( array() )
  call initJquery( array("","","072014") )
  call initJqueryUi( array("","","","072014") )
  call initBootstrap( array("072015","","home") )
%>  	

</head>

<body>
<input type="file" id="my_file_input" />
<!-- <button  type="button" class="btn btn-primary" id="model2"> Open Model #2 </button> -->
<script src="js-xlsx-master/dist/xlsx.full.min.js"></script>
<script type="text/javascript" src="sdn-import-excel.js"></script>
<script>

  $(function() {

    XLStoJSON.init({
    	el:'#my_file_input',
    	url:'api.asp',
    	import:'importboat,someimport,someimport1,someimport2'
    })
    
  });

</script>

</body>
</html>
<!-- #include file="../site/deconnect.asp" -->

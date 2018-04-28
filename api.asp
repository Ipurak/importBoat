<!-- #include file="../style/AVEImgClass.asp"-->
<!-- #include file="../search/connect.asp"-->
<!-- #include file="../site/lang.asp" -->
<!-- #include file="../style/Functions.asp" -->
<!-- #include file="../style/class.json.asp" -->
<!-- #include file="../style/class.json.util.asp" -->
<%
'Author: Ham
'Created: 06/12/2017
'########## [START] HOW TO SETTING ##########'
' The importboat.data in this object need to match each columns in Excel file
' Example: Model is column A have to set object like below
' {"position":"A", "name":"Model", "index":1, "map":null }

'importboat.name >> Name of Import type, it shown on menu
'importboat.headerRange >> It's How many rows of title in Excel file format
'importboat.icon >> Icon class
'########## [END] HOW TO SETTING ##########'

Dim DicJson
Set DicJson=Server.CreateObject("Scripting.Dictionary")
DicJson.Add "importboat","{"&_
"      ""name"":""Import Boat"","&_
"      ""headerRange"":2,"&_
"      ""icon"":""fa fa-ship"","&_
"      ""data"":{"&_
"      ""Model"":{""position"":""A"", ""name"":""Model"", ""index"":1, ""map"":null },"&_
"      ""Name"":{""position"":""B"", ""name"":""Name"", ""index"":2, ""map"":null },"&_
"      ""Boatbuilder"":{""position"":""C"", ""name"":""Boat builder"", ""index"":3, ""map"":null },"&_
"      ""Type"":{""position"":""D"", ""name"":""Type"", ""index"":4, ""map"":null },"&_
"      ""Caution"":{""position"":""E"", ""name"":""Caution"", ""index"":5, ""map"":null },"&_
"      ""LOAm"":{""position"":""F"", ""name"":""L.O.A (m)"", ""index"":6, ""map"":null },"&_
"      ""yearbuilt"":{""position"":""G"", ""name"":""year built"", ""index"":7, ""map"":null },"&_
"      ""MaxPax"":{""position"":""H"", ""name"":""Max Pax"", ""index"":8, ""map"":null },"&_
"      ""DoubleCabins"":{""position"":""I"", ""name"":""Double Cabins"", ""index"":9, ""map"":null },"&_
"      ""SingleCabins"":{""position"":""J"", ""name"":""Single Cabins"", ""index"":10, ""map"":null },"&_
"      ""Noofbathrooms"":{""position"":""K"", ""name"":""no. of bathrooms"", ""index"":11, ""map"":null },"&_
"      ""Toilets"":{""position"":""L"", ""name"":""toilets"", ""index"":12, ""map"":null },"&_
"      ""Beam"":{""position"":""M"", ""name"":""Beam (m)"", ""index"":13, ""map"":null },"&_
"      ""Shower"":{""position"":""N"", ""name"":""shower"", ""index"":14, ""map"":null },"&_
"      ""HorsePower"":{""position"":""O"", ""name"":""Horse power (engine)"", ""index"":15, ""map"":null },"&_
"      ""FuelCapacity"":{""position"":""P"", ""name"":""Fuel capacity (l)"", ""index"":16, ""map"":null },"&_
"      ""WaterCapacity"":{""position"":""Q"", ""name"":""Water capacity (l)"", ""index"":17, ""map"":null },"&_
"      ""HotWater"":{""position"":""R"", ""name"":""Hot water"", ""index"":18, ""map"":null },"&_
"      ""SailSurfaceArea"":{""position"":""S"", ""name"":""sail surface area (m2)"", ""index"":19, ""map"":null },"&_
"      ""English"":{""position"":""T"", ""name"":""English"", ""index"":20, ""map"":null },"&_
"      ""French"":{""position"":""U"", ""name"":""French"", ""index"":21, ""map"":null },"&_
"      ""Italia"":{""position"":""V"", ""name"":""Italia"", ""index"":22, ""map"":null },"&_
"      ""Spanish"":{""position"":""W"", ""name"":""Spanish"", ""index"":23, ""map"":null }"&_
"      }"&_
"}"
DicJson.Add "someimport1","{"&_
"      ""name"":""Some Import Some Import"","&_
"      ""headerRange"":2,"&_
"      ""icon"":""fa fa-anchor"","&_
"      ""data"":{"&_
"      ""Shower"":{""position"":""N"", ""name"":""shower"", ""index"":14, ""map"":null },"&_
"      ""HorsePower"":{""position"":""O"", ""name"":""Horse power (engine)"", ""index"":15, ""map"":null },"&_
"      ""FuelCapacity"":{""position"":""P"", ""name"":""Fuel capacity (l)"", ""index"":16, ""map"":null },"&_
"      ""WaterCapacity"":{""position"":""Q"", ""name"":""Water capacity (l)"", ""index"":17, ""map"":null },"&_
"      ""HotWater"":{""position"":""R"", ""name"":""Hot water"", ""index"":18, ""map"":null },"&_
"      ""SailSurfaceArea"":{""position"":""S"", ""name"":""sail surface area (m2)"", ""index"":19, ""map"":null },"&_
"      ""English"":{""position"":""T"", ""name"":""English"", ""index"":20, ""map"":null },"&_
"      ""French"":{""position"":""U"", ""name"":""French"", ""index"":21, ""map"":null },"&_
"      ""Italia"":{""position"":""V"", ""name"":""Italia"", ""index"":22, ""map"":null },"&_
"      ""Spanish"":{""position"":""W"", ""name"":""Spanish"", ""index"":23, ""map"":null }"&_
"      }"&_
"    }"
DicJson.Add "someimport2","{"&_
"      ""name"":""Some Import Some Import"","&_
"      ""headerRange"":2,"&_
"      ""icon"":""fa fa-flask"","&_
"      ""data"":{"&_
"      ""Shower"":{""position"":""N"", ""name"":""shower"", ""index"":14, ""map"":null },"&_
"      ""HorsePower"":{""position"":""O"", ""name"":""Horse power (engine)"", ""index"":15, ""map"":null },"&_
"      ""FuelCapacity"":{""position"":""P"", ""name"":""Fuel capacity (l)"", ""index"":16, ""map"":null },"&_
"      ""WaterCapacity"":{""position"":""Q"", ""name"":""Water capacity (l)"", ""index"":17, ""map"":null },"&_
"      ""HotWater"":{""position"":""R"", ""name"":""Hot water"", ""index"":18, ""map"":null },"&_
"      ""SailSurfaceArea"":{""position"":""S"", ""name"":""sail surface area (m2)"", ""index"":19, ""map"":null },"&_
"      ""English"":{""position"":""T"", ""name"":""English"", ""index"":20, ""map"":null },"&_
"      ""French"":{""position"":""U"", ""name"":""French"", ""index"":21, ""map"":null },"&_
"      ""Italia"":{""position"":""V"", ""name"":""Italia"", ""index"":22, ""map"":null },"&_
"      ""Spanish"":{""position"":""W"", ""name"":""Spanish"", ""index"":23, ""map"":null }"&_
"      }"&_
"    }"
DicJson.Add "someimport3","{"&_
"      ""name"":""Some Import Some Import"","&_
"      ""headerRange"":2,"&_
"      ""icon"":""fa fa-flask"","&_
"      ""data"":{"&_
"      ""Shower"":{""position"":""N"", ""name"":""shower"", ""index"":14, ""map"":null },"&_
"      ""HorsePower"":{""position"":""O"", ""name"":""Horse power (engine)"", ""index"":15, ""map"":null },"&_
"      ""FuelCapacity"":{""position"":""P"", ""name"":""Fuel capacity (l)"", ""index"":16, ""map"":null },"&_
"      ""WaterCapacity"":{""position"":""Q"", ""name"":""Water capacity (l)"", ""index"":17, ""map"":null },"&_
"      ""HotWater"":{""position"":""R"", ""name"":""Hot water"", ""index"":18, ""map"":null },"&_
"      ""SailSurfaceArea"":{""position"":""S"", ""name"":""sail surface area (m2)"", ""index"":19, ""map"":null },"&_
"      ""English"":{""position"":""T"", ""name"":""English"", ""index"":20, ""map"":null },"&_
"      ""French"":{""position"":""U"", ""name"":""French"", ""index"":21, ""map"":null },"&_
"      ""Italia"":{""position"":""V"", ""name"":""Italia"", ""index"":22, ""map"":null },"&_
"      ""Spanish"":{""position"":""W"", ""name"":""Spanish"", ""index"":23, ""map"":null }"&_
"      }"&_
"    }"

	typeApi      = request("type")
	importOption = request("import")
	Json 		     = request("json")
	if typeApi = "getinit" then
		getinit()
	elseif typeApi ="import" then
		' import()
	end if

	function getinit()
		initJson     = ""
		dataJson  	 = ""
		importArr    = split(importOption,",")
		importLength = ubound(importArr)

		Set	initObject = jsObject()
		for i = 0 to importLength
			if DicJson.Item(importArr(i))&"a" <> "a" then
				initObject(importArr(i)) = DicJson.Item(importArr(i))
			else
				initObject(importArr(i)) = null	
			end if
		next 
		initObject.Flush
	end function

	function import()
		'Set the command
		DIM cmd
		SET cmd = Server.CreateObject("ADODB.Command")
		SET cmd.ActiveConnection = cn

		'Prepare the stored procedure
		cmd.CommandText = "HamTest1"
		cmd.CommandType = 4  'adCmdStoredProc

		cmd.Parameters.Refresh
		cmd.Parameters("@JSONText") = Json

		'Execute the stored procedure
		'This returns recordset but you dont need it
		cmd.Execute
	end function


%>
<!-- #include file="../site/deconnect.asp" -->

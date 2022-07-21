'Apertura de navegador 
Set objBrowser = Browser("maps - Google Maps")
Set objPage = objBrowser.Page("maps - Google Maps") 
'Busqueda de Thales desde la data
strLocations = Datatable.Value("THALES")
objPage.WebEdit("a").Set strLocations
objPage.WebButton("Buscar").Click
wait 4

' Mapeo de campos basicos de thales 
Browser("maps - Google Maps").Page("Thales International Suc._2").WebElement("17442442").CheckPoint("17442442")
Browser("maps - Google Maps").Page("Thales International Suc._2").WebElement("Cra. 12 #93-8, Bogotá").CheckPoint.("Cra. 12 #93-8, Bogotá")
Browser("maps - Google Maps").Page("Thales International Suc._2").WebElement("thalesgroup.com").CheckPoint.("thalesgroup.com")

' Inicio de ruta para la primera y ultima  ubicacion definida 
Browser("maps - Google Maps").Page("Thales International Suc.").WebElement("Cómo llegar").Click
Set objPage = objBrowser.Page("maps - Google Maps") 
strLocations = Datatable.Value("LOCATIONS")
objPage.WebEdit("a").Set strLocations
objPage.WebButton("Buscar").Click
Browser("maps - Google Maps").Page("de Modelia, ZONA 5, Bogotá").WebElement("18 min")
wait 4
Browser("maps - Google Maps").Page("Thales International Suc.").WebElement("Cómo llegar").Click
Set objPage = objBrowser.Page("maps - Google Maps") 
strLocations = Datatable.Value("LOCATIONS")
objPage.WebEdit("a").Set strLocations
objPage.WebButton("Buscar").Click
Browser("maps - Google Maps").Page("de EFECTY WORLD TRADE").WebElement("5 min").Click
wait 4

Browser("maps - Google Maps").Page("Thales International Suc.").WebElement("Cómo llegar").Click
Set objPage = objBrowser.Page("maps - Google Maps") 
strLocations = Datatable.Value("LOCATIONS")
objPage.WebEdit("a").Set strLocations
objPage.WebButton("Buscar").Click
Browser("maps - Google Maps").Page("de Kennedy, Bogotá a Thales").WebElement("21 min").Click
wait 4

Browser("maps - Google Maps").Page("Thales International Suc.").WebElement("Cómo llegar").Click
Set objPage = objBrowser.Page("maps - Google Maps") 
strLocations = Datatable.Value("LOCATIONS")
objPage.WebEdit("a").Set strLocations
objPage.WebButton("Buscar").Click
Browser("maps - Google Maps").Page("de Bosa, Bogotá a Thales").WebElement("28 min ").Click
wait 4

Browser("maps - Google Maps").Page("Thales International Suc.").WebElement("Cómo llegar").Click
Set objPage = objBrowser.Page("maps - Google Maps") 
strLocations = Datatable.Value("LOCATIONS")
objPage.WebEdit("a").Set strLocations
objPage.WebButton("Buscar").Click
Browser("maps - Google Maps").Page("de Soacha, Cundinamarca").WebElement("32 min").Click

wait 4






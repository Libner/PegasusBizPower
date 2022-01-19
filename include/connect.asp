<%set con=server.createobject("adodb.connection")
	  con.open Application("ConnectionString")
	    
	  font_face        ="arial" ' font face for texts(optional, default - "arial")
	  pxpt             ="pt"    ' metrical(optional, default - "pt")
	  const PageWidth  = "100%"
	  
'added by Mila 16/07/2017 - DB names for using in joined queries between two databases
'!!!!Change in PRODUCTION!!!!!!!!!!!!!!!
pegasusDBName=Application("pegasusDBName")
bizpowerDBName=Application("bizpowerDBName")
	  %>
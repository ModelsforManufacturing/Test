Dim innovatorServer 
Dim innovatorServer1
innovatorServer = "http://vicente.us.es/InnovatorServer/Server/InnovatorServer.aspx" 
innovatorServer1="http://vicente.us.es/innovatorServer/vault/vaultserver.aspx"
Dim database: database = "PracticaGrupo15" 
Dim loginName: loginName = "admin" 
Dim password: password = "607920b64fe136f9ab2389e371852af2"  'MD5 hash of the password 
Dim soapStart: soapStart = "<SOAP-ENV:Envelope xmlns:SOAP-ENV='http://schemas.xmlsoap.org/soap/envelope/' " &_ 
    "encodingStyle='http://schemas.xmlsoap.org/soap/encoding/'><SOAP-ENV:Body>" 
Dim soapEnd: soapEnd = "</SOAP-ENV:Body></SOAP-ENV:Envelope>" 
Dim soapAction: soapAction = "ApplyAML" 
Dim body: 'body = "<Item type='Part' action='get' select='item_number'/>" 
'body="<AML><Item type=""Part"" action=""delete"" where =""item_number='Prueba 1'""></Item></AML>"
dim ts
ts=readBinary ("c:\datos\ejemplo.txt")
Wscript.Echo "fichero: " & ts
Dim body1:
body1="<AML><Item type='Document' action='add'>"
body1=body1 &"<item_number>myNumber</item_number><name>myName</name><description>myDescription</description></Item></ApplyAML>"

body1="<AML>" & _
"<Item type=""Part BOM"" action=""delete"" where=""keyed_name='Part 1.2' and source_id='Prueba 1'"">" & _	
	"</Item>" & _
"</AML>"    
'10BECF6600F845918E17DF87A110DC2E
body1="<AML><Item type=""Part"" action=""get"" >" & _
  "<item_number>Prueba 1</item_number>" & _
  "<Relationships>" & _    
  "<Item type=""Part BOM"" action=""get"" >" & _	
	"</Item>" & _              
  "</Relationships>" & _
"</Item>" & _
"</AML>"
Dim soap 
soap = soapStart & "<" & SOAPAction & " xmlns:m='http://www.aras-corp.com/'>" &_ 
           body & "</" & soapAction & ">" & soapEnd 
		   Wscript.Echo soap
dim bounday
boundary= "————————-BRdnIy5hBONlyI"
Dim content1
content = "–"+boundary &_ 
		 "Content-Disposition: form-data; name=""SOAPACTION"";" &_
		 "Content-Type: text/plain" &_
		 "ApplyItem"
Dim sFile
sFile=".\CleanSky.eap"
inByteArray=readbytes(sFile)
base64Encoded=encodeBase64(inByteArray)

content = "–"+boundary +" &_
"Content-Disposition: form-data; name=""SOAPACTION"";" &_
"Content-Type: text/plain" &_
"ApplyItem"

Dim http: Set http = CreateObject("Msxml2.ServerXMLHTTP") 
http.open "POST", innovatorServer, false 
content = "–"+boundary +" &_
"Content-Disposition: form-data; name=""XMLDATA""" &_
"Content-Type: text/plain" 
        
        body = "   <Item type='Document' id='76126C1815DA441B80E11492ACF437A4' action='add'>";
        body = body & "    <item_number>X-1001</item_number>";
        body = body & "    <name>Test</name>";
        body = body & "    <locked_by_id keyed_name='admin'>30B991F927274FA3829655F50C99472E</locked_by_id>";
        body = body & "    <Relationships>";
        body = body & "     <Item type='Document File' id='A311B7E4A1414BF9971F01867E919A43' action='add'>";
        body = body & "      <related_id>";
        body = body & "       <Item type='File' id='534007B45DBA48249AC49C8F90A26DF4' action='add'>";
        body = body & "        <filename>CleanSky.eap</filename>";
        body = body & "        <checkedout_path>.\</checkedout_path>";
        body = body & "        <new_version>1</new_version>";
        body = body & "        <file_size>11</file_size>";
        body = body & "        <checksum>FB53A94DDC6855BA4DCB9E9BD10E0AC0</checksum>";
        body = body & "        <Relationships>";
        body = body & "         <Item type='Located' id='78FF03607AAA4926B80E5CE628729AF5' action='add'>";
        body = body & "          <source_id keyed_name='testdoc.txt'>534007B45DBA48249AC49C8F90A26DF4</source_id>";
        body = body &"          <related_id keyed_name='Default'>67BBB9204FE84A8981ED8313049BA06C</related_id>";
        body = body & "         </Item>";
        body = body & "        </Relationships>";
        body = body & "       </Item>";
        body = body & "      </related_id>";
        body = body & "      <source_id keyed_name='X-1001'>76126C1815DA441B80E11492ACF437A4</source_id>";
        body = body & "     </Item>"    ;
        body = body & "    </Relationships>";
        body = body & "   </Item>";

http.setRequestHeader "SOAPaction", soapAction 
http.setRequestHeader "Content-Type", "multipart/form-data; boundary="& boundary
http.setRequestHeader "Content-Length", Len(request)

http.setRequestHeader "AUTHUSER",  loginName 
http.setRequestHeader "AUTHPASSWORD", password 
http.setRequestHeader "DATABASE", database 
http.send(soap) 
Dim response: response = http.responseText 
Dim responseDom: Set responseDom  = CreateObject("microsoft.xmldom") 
responseDom.loadXML(response) 
Dim userItems: Set userItems = responseDom.selectNodes("//Item[@type='User']") 
Wscript.Echo "Number of users: " & userItems.Length 
Wscript.Echo response
soapAction = "logoff" 
body = "logoff" 
soap = soapStart & "<" & SOAPAction & " xmlns:m='http://www.aras-corp.com/'>" &_ 
           body & "</" & soapAction & ">" & soapEnd 
http.open "POST", InnovatorServer, false 
http.setRequestHeader "SOAPaction", soapAction 
http.setRequestHeader "AUTHUSER",  loginName
http.setRequestHeader "AUTHPASSWORD", password 
http.setRequestHeader "DATABASE", database 
http.send(soap)


Private function readBytes(file)
  dim inStream
  ' ADODB stream object used
  set inStream = WScript.CreateObject("ADODB.Stream")
  ' open with no arguments makes the stream an empty container 
  inStream.Open
  inStream.type= 1
  inStream.LoadFromFile(file)
  readBytes = inStream.Read()
end function

Private function encodeBase64(bytes)
  dim DM, EL
  Set DM = CreateObject("Microsoft.XMLDOM")
  ' Create temporary node with Base64 data type
  Set EL = DM.createElement("tmp")
  EL.DataType = "bin.base64"
  ' Set bytes, get encoded String
  EL.NodeTypedValue = bytes
  encodeBase64 = EL.Text
end function
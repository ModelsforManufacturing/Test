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
Dim body: body = "<Item type='Part' action='get' select='item_number'/>" 

dim bounday
boundary= "————————-BRdnIy5hBONlyI"

content = "–"+boundary &_ 
		 "Content-Disposition: form-data; name=""SOAPACTION"";" &_
		 "Content-Type: text/plain" &_
		 "ApplyItem"
	
Dim soap 
soap = content & soapStart & "<" & SOAPAction & " xmlns:m='http://www.aras-corp.com/'>" &_ 
           body & "</" & soapAction & ">" & soapEnd & content
		   'Wscript.Echo soap
Dim content1

Dim sFile
sFile=".\ejemplo.vbs"
inByteArray=readbytes(sFile)
base64Encoded=encodeBase64(inByteArray)
Wscript.Echo SimpleBinaryToString(base64Encoded)
content = "–" & boundary &_
"Content-Disposition: form-data; name=""SOAPACTION"";" &_
"Content-Type: text/plain" &_
"ApplyItem"

Dim http: Set http = CreateObject("Msxml2.ServerXMLHTTP") 
http.open "POST", innovatorServer1, false 
http.setRequestHeader "Content-Type", "text/xml; charset=UTF-8"
http.setRequestHeader "SOAPAction", "BeginTransaction"

http.setRequestHeader "AUTHUSER",  loginName 
http.setRequestHeader "AUTHPASSWORD", password 
http.setRequestHeader "DATABASE", database 
http.send("SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" ><SOAP-ENV:Body><BeginTransaction></BeginTransaction></SOAP-ENV:Body></SOAP-ENV:Envelope>") 

Dim response: response = http.responseText 
Dim responseDom: Set responseDom  = CreateObject("microsoft.xmldom") 
responseDom.loadXML(response) 
Dim colNodes: Set colNodes = responseDom.selectNodes ("//BeginTransactionResponse/Result") 
dim tran
For Each objNode in colNodes
  Wscript.Echo objNode.Text 
  tran=objNode.Text
Next
innovatorServer1="http://vicente.us.es/innovatorserver/vault/vaultserver.aspx"
Dim http1: Set http1 = CreateObject("Msxml2.ServerXMLHTTP") 
http1.open "POST", innovatorServer1, false 
http1.setRequestHeader "Content-Type", "application/octet-stream"
http1.setRequestHeader "Content-Disposition", "attachment; filename*=utf-8''ejemplo.vbs"
http1.setRequestHeader "transactionid", tran
http1.setRequestHeader "Content-Length", "5946"
http1.setRequestHeader "SOAPAction", "UploadFile"

http1.setRequestHeader "AUTHUSER",  loginName 
http1.setRequestHeader "AUTHPASSWORD", password 
http1.setRequestHeader "DATABASE", database 

http1.send(base64Encoded) 
Dim response1: response1 = http1.responseText 

Wscript.Echo response1
'http.send(soap)


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
Function SimpleBinaryToString(Binary)
  'SimpleBinaryToString converts binary data (VT_UI1 | VT_ARRAY Or MultiByte string)
  'to a string (BSTR) using MultiByte VBS functions
  Dim I, S
  For I = 1 To LenB(Binary)
    S = S & Chr(AscB(MidB(Binary, I, 1)))
  Next
  SimpleBinaryToString = S
End Function


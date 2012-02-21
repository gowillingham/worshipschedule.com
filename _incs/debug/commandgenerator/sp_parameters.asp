<%@ EnableSessionState=False Language=VBScript %>
<%
'Response.Buffer=true
Response.Expires = 0
'The code can be coopied and pasted into the server-side script code
'Querystring variables:
'	Srv - sql server name
'	DB - database name
'	Proc - stored procedure name
' User - SQL Login Name (code can be rewritten to use integrated security
' pwd - password
dim cnSQL, cmd, params, sProcName, sDBName, sSrv
dim param, xmlDataTypes, sADO, oRoot, bStatus,xPE
dim oNode, sParamDirection, sDataType,sUser, sPWD,sMsg

const csCmdVar = "cmd"
const csCnVar = "cnn"

sMsg = "The following querystring parameters must be supplied:<BR>Srv - Server name<BR>DB - database name<BR>Proc - Stored Procedure name, qualified with owner name if not dbo<BR>User - SQL User Name<BR>pwd - Password"
sMsg = sMsg & "<BR>Example: http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?srv=servername&db=test&proc=procedurename&User=username&pwd=password"
sProcName = Request("Proc")
if len(sProcName) = 0 then 
	Response.Write "No procedure name was supplied<BR>"
	Response.Write sMsg 
	Response.End
end if

sSrv = Request("Srv")
if len(sSrv) = 0 then 
	Response.Write "No server name was supplied<BR>"
	Response.Write sMsg
	Response.End
end if

sDBName = Request("db")
if len(sDBName) = 0 then 
	Response.Write "No database name was supplied<BR>"
	Response.Write sMsg
	Response.End
end if

sUser = Request("User")
if len(sUser) = 0 then 
	Response.Write "No user name was supplied<BR>"
	Response.Write sMsg
	Response.End
end if

sPWD = Request("pwd")
if len(sPWD) = 0 then 
	Response.Write "No password was supplied<BR>"
	Response.Write sMsg
	Response.End
end if

sADO = "<datatypes ><datatype name=""adEmpty"" value=""0"" /><datatype name=""adTinyInt"" value=""16"" /><datatype name=""adSmallInt"" value=""2"" /><datatype name=""adInteger"" value=""3"" /><datatype name=""adBigInt"" value=""20"" /><datatype name=""adUnsignedTinyInt"" value=""17"" /><datatype name=""adUnsignedSmallInt"" value=""18"" /><datatype name=""adUnsignedInt"" value=""19"" /><datatype name=""adUnsignedBigInt"" value=""21"" /><datatype name=""adSingle"" value=""4"" /><datatype name=""adDouble"" value=""5"" /><datatype name=""adCurrency"" value=""6"" /><datatype name=""adDecimal"" value=""14"" /><datatype name=""adNumeric"" value=""131"" /><datatype name=""adBoolean"" value=""11"" /><datatype name=""adError"" value=""10"" /><datatype name=""adUserDefined"" value=""132"" /><datatype name=""adVariant"" value=""12"" /><datatype name=""adIDispatch"" value=""9"" /><datatype name=""adIUnknown"" value=""13"" /><datatype name=""adGUID"" value=""72"" /><datatype name=""adDate"" value=""7"" /><datatype name=""adDBDate"" value=""133"" /><datatype name=""adDBTime"" value=""134"" /><datatype name=""adDBTimeStamp"" value=""135"" /><datatype name=""adBSTR"" value=""8"" /><datatype name=""adChar"" value=""129"" /><datatype name=""adVarChar"" value=""200"" /><datatype name=""adLongVarChar"" value=""201"" /><datatype name=""adWChar"" value=""130"" /><datatype name=""adVarWChar"" value=""202"" /><datatype name=""adLongVarWChar"" value=""203"" /><datatype name=""adBinary"" value=""128"" /><datatype name=""adVarBinary"" value=""204"" /><datatype name=""adLongVarBinary"" value=""205"" /><datatype name=""adChapter"" value=""136"" /><datatype name=""adFileTime"" value=""64"" /><datatype name=""adDBFileTime"" value=""137"" /><datatype name=""adPropVariant"" value=""138"" /><datatype name=""adVarNumeric"" value=""139"" /></datatypes>"
Set xmlDataTypes=Server.CreateObject("msxml2.FreeThreadedDOMDocument")
xmlDataTypes.async = false
bStatus= xmlDataTypes.loadXML(sADO)
if bStatus = false then
	Set xPE = xmlDataTypes.parseError
	strMessage = "errorCode = " & xPE.errorCode & "<BR>"
	strMessage = strMessage & "reason = " & xPE.reason & vbCrLf
	strMessage = strMessage & "Line  = " & xPE.Line & vbCrLf
	strMessage = strMessage & "linepos = " & xPE.linepos & "<BR>"
	strMessage = strMessage & "filepos = " & xPE.filepos & "<BR>"
	strMessage = strMessage & "srcText = " & xPE.srcText & "<BR>"
	Response.Write strMessage
	Response.End
else	
'	Set oRoot = xmlDataTypes.documentelement
'	xmlDataTypes.insertbefore xmlDataTypes.createprocessinginstruction("xml", " version=""1.0"""),oRoot
	'Response.Write xmlDataTypes.xml
	'Response.End
end if



sConnect="Provider=SQLOLEDB.1;Password=" & sPWD & ";Persist Security Info=False;User ID=" & sUser & ";Initial Catalog=" & sDBName & ";Data Source=" & Request("Srv") & ";Application Name=ProcParams"
'Response.Write "sConnnect contains " & sConnect
'Response.end
Set cnSQL = server.CreateObject("ADODB.Connection")
cnSQL.ConnectionString=sConnect
cnSQL.Open


Set cmd=server.CreateObject("ADODB.Command")
cmd.CommandType=adcmdstoredproc
cmd.CommandText = sProcName
cmd.ActiveConnection=cnSQL
Set params = cmd.Parameters
params.refresh

Response.Write "Dim " & csCmdVar & ", " & csCnVar & ", parm<BR>"
Response.Write "Set " & csCmdVar & " = Server.CreateObject(""ADODB.Command"")<BR><br />"
Response.Write "With " & csCmdVar & "<BR>"
Response.Write ".CommandType = adCmdStoredProc<BR>"
Response.Write ".CommandText = """ & sProcName & """<BR>"
Response.Write "Set cnn = Server.CreateObject(""ADODB.Connection"")<br />"
Response.Write "cnn.Open Application.Value(""CNN_STR"")<br />"
Response.Write ".ActiveConnection = cnn<BR>"
Response.Write "End With<br /><br />"
for each param in params		
		Response.Write "cmd.Parameters.Append cmd.CreateParameter(""" & param.name & """, "
		sDataType=param.type	
		Set oNode=xmlDataTypes.selectSingleNode("/datatypes/datatype[@value=" & sDataType & "]")
		sDataType = oNode.getattribute("name")
		Response.Write sDataType & ", "
		select case param.direction
			case 0: Response.Write "adParamUnknown, " & param.size 
			case 1: Response.Write "adParamInput, " & param.size 
			case 2: Response.Write "adParamOutput, " & param.size
			case 3: Response.Write "adParamInputOutput, " & param.size 
			case 4: Response.Write "adParamReturnValue, " & param.size
		end select
		if param.direction=adParamReturnValue then
			Response.Write ")<BR>"
		else
			Response.Write ", )<BR>"	
		end if	
		if instr("adDecimal,adNumeric",sDataType) > 0 then
			Response.Write "parm.Precision=" & param.precision & "<BR>"
			Response.Write "parm.NumericScale=" & param.numericscale & "<BR>"		
		end if	
		
next
Response.Write "<br />cmd.Execute ,,adExecuteNoRecords<BR><br />"
Response.Write "Set parm = Nothing: Set cmd = Nothing<br />"
Response.Write "cnn.Close: Set cnn = Nothing<br />"

Response.Write "<BR>If you do not wish to supply values for your output parameters, replace adParamInputOutput with adParamOutput"
Set xmlDataTypes=nothing
Set cmd=nothing
cnSQL.Close
Set cnSQL = nothing



%>
<%
Option Explicit

Dim userSettings		: Set userSettings = New cUserSettings

Dim key					: key = Request.Form("key")
Dim formValue			: formValue= Request.Form("value")
Dim id					: id = request.Form("id")
Dim val

' massage form value, all values are bit 0/1
If CBool(formValue) Then
	val = 1
Else
	val = 0
End If

userSettings.Load(id)
userSettings.ChangeSetting key, val
userSettings.Save id, ""

%>

<!--#INCLUDE VIRTUAL="/_incs/class/user_settings_cls.asp"-->

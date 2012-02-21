<script language="vbscript" type="text/vbscript" runat="server">
Class clsRC4
	Private mStrKey
	Private mBytKeyAry(255)
	Private mBytCypherAry(255)
	
	Private Sub InitializeCypher()
		
		Dim lBytJump
		Dim lBytIndex
		Dim lBytTemp
	
		For lBytIndex = 0 To 255
		mBytCypherAry(lBytIndex) = lBytIndex
		Next
		' Switch values of Cypher arround based off of index and Key value
		lBytJump = 0
		For lBytIndex = 0 To 255
	
			' Figure index To switch
		lBytJump = (lBytJump + mBytCypherAry(lBytIndex) + mBytKeyAry(lBytIndex)) Mod 256
		
		' Do the switch
		lBytTemp					= mBytCypherAry(lBytIndex)
		mBytCypherAry(lBytIndex)	= mBytCypherAry(lBytJump)
		mBytCypherAry(lBytJump)		= lBytTemp
		
		Next
	End Sub
	
	Public Property Let Key(ByRef pStrKey)
		Dim lLngKeyLength
		Dim lLngIndex
		
		if pStrKey = mStrKey Then Exit Property
		lLngKeyLength = Len(pStrKey)
		if lLngKeyLength = 0 Then Exit Property
		mStrKey = pStrKey
		lLngKeyLength = Len(pStrKey)
		For lLngIndex = 0 To 255
		mBytKeyAry(lLngIndex) = Asc(Mid(pStrKey, ((lLngIndex) Mod (lLngKeyLength)) + 1, 1))
		Next
	End Property
	
	Public Property Get Key()
		Key = mStrKey
	End Property
	
	Public function Crypt(ByRef pStrMessage)
		Dim lBytIndex
		Dim lBytJump
		Dim lBytTemp
		Dim lBytY
		Dim lLngT
		Dim lLngX
		
		' Validate data
		if Len(mStrKey & "") = 0 Then Exit function
		if Len(pStrMessage & "") = 0 Then Exit function
		Call InitializeCypher()
		
		lBytIndex = 0
		lBytJump = 0
		For lLngX = 1 To Len(pStrMessage)
		lBytIndex = (lBytIndex + 1) Mod 256 ' wrap index
		lBytJump = (lBytJump + mBytCypherAry(lBytIndex)) Mod 256 ' wrap J+S()
		
			' Add/Wrap those two	
		lLngT = (mBytCypherAry(lBytIndex) + mBytCypherAry(lBytJump)) Mod 256
		
		' Switcheroo
		lBytTemp					= mBytCypherAry(lBytIndex)
		mBytCypherAry(lBytIndex)	= mBytCypherAry(lBytJump)
		mBytCypherAry(lBytJump)		= lBytTemp
		lBytY = mBytCypherAry(lLngT)
			' Character Encryption ...
		Crypt = Crypt & Chr(Asc(Mid(pStrMessage, lLngX, 1)) Xor lBytY)
		Next
		
	End function
End Class
'**************************************
' Name: RC4 Class
' Description:Applys Encryption/Decrypti
'     on to strings. I think just about everyo
'     ne who has seen my code knows how I love
'     classes. This version is more "cleaned u
'     p" and thrown into a nice little class f
'     or an object oriented feeling. (If only 
'     ASP was object oriented, I would be a ha
'     ppy camper). It is also a little more op
'     timized to run quicker if you change the
'     Key/Password often.
' By: Lewis Moten
'
'	  This code is copyrighted and has    
'	  limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/vb/scripts/Sho
'     wCode.asp?txtCodeId=6649&lngWId=4    
'	  for details.    
'**************************************
</script>

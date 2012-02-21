<script language="vbscript" type="text/vbscript" runat="server">

Class cSubscription
	Private m_subscriptionID
	Private m_subscriptionName
	Private m_description
	Private m_termLength
	Private m_price
	Private m_isEnabled
	
	Private CLASS_NAME
	Private TYPE_OF
	
	Public Property Let subscriptionID(val)
		m_subscriptionID = val
	End Property
	
	Public Property Get subscriptionID()
		subscriptionID = m_subscriptionID
	End Property

	Public Property Let subscriptionName(val)
		m_subscriptionName = val
	End Property
	
	Public Property Get subscriptionName()
		subscriptionName = m_subscriptionName
	End Property

	Public Property Let description(val)
		m_description = val
	End Property
	
	Public Property Get description()
		description = m_description
	End Property

	Public Property Let termLength(val)
		m_termLength = val
	End Property
	
	Public Property Get termLength()
		termLength = m_termLength
	End Property

	Public Property Let price(val)
		m_price = val
	End Property
	
	Public Property Get price()
		price = m_price
	End Property

	Public Property Let isEnabled(val)
		m_isEnabled = val
	End Property
	
	Public Property Get isEnabled()
		isEnabled = m_isEnabled
	End Property

	Public Property Get IsTypeOf()
		IsTypeOf = TYPE_OF
	End Property
	
	Public Function List()
		Dim cnn				: Set cnn = Server.CreateObject("ADODB.Connection")
		Dim rs				: Set rs = Server.CreateObject("ADODB.Recordset")
		
		' 0-SubscriptionID 1-Name 2-Desc 3-TermLength 4-Price 5-IsEnabled

		Call cnn.Open(Application.Value("CNN_STR"))
		cnn.up_clientGetSubscriptionList rs
		If Not rs.EOF Then list = rs.GetRows()
		
		If rs.State = adStateOpen Then rs.Close(): Set rs = Nothing
		Set cnn = Nothing
	End Function
	
	Private Sub Class_Initialize()
		CLASS_NAME = "cSubscription"
		TYPE_OF = "ws.Subscription"
	End Sub
	
	Private Sub Class_Terminate()
	
	End Sub
End Class
</script>

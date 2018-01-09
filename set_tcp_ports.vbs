set wmiComputer = GetObject( _
    "winmgmts:" _
    & "\\.\root\Microsoft\SqlServer\ComputerManagement13")
set tcpProperties = wmiComputer.ExecQuery( _
    "select * from ServerNetworkProtocolProperty " _
    & "where InstanceName='SQLEXPRESS' and " _
    & "(ProtocolName='Tcp' and IPAddressName='IP1') or" _
	& "(ProtocolName='Tcp' and IPAddressName='IP2') or" _
	& "(ProtocolName='Tcp' and IPAddressName='IP3') or" _
	& "(ProtocolName='Tcp' and IPAddressName='IP4') or" _ 
	& "(ProtocolName='Tcp' and IPAddressName='IP5') or" _ 
	& "(ProtocolName='Tcp' and IPAddressName='IPAll')")

for each tcpProperty in tcpProperties
    dim requestedValue

    if tcpProperty.PropertyName = "TcpPort" then
		portnum = "1433"
		tcpProperty.SetStringValue(portnum)
	end if
next
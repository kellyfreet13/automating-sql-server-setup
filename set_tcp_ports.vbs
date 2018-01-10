set wmiComputer = GetObject( _
    "winmgmts:" _
    & "\\.\root\Microsoft\SqlServer\ComputerManagement13")
set tcpProperties = wmiComputer.ExecQuery( _
    "SELECT * FROM ServerNetworkProtocolProperty " _
    & "WHERE InstanceName='SQLEXPRESS' AND " _
    & "ProtocolName='Tcp' AND (IPAddressName LIKE 'IP%')")
for each tcpProperty in tcpProperties
    dim requestedValue

    if tcpProperty.PropertyName = "TcpPort" then
		portnum = "1433"
		tcpProperty.SetStringValue(portnum)
	end if
next
set wmiComputer = GetObject( _
    "winmgmts:" _
    & "\\.\root\Microsoft\SqlServer\ComputerManagement13")
set tcpProperties = wmiComputer.ExecQuery( _
    "select * from ServerNetworkProtocolProperty " _
    & "where InstanceName='SQLEXPRESS' and " _
    & "(ProtocolName='Tcp' and IPAddressName='IP%')"

for each tcpProperty in tcpProperties
    dim requestedValue

    if tcpProperty.PropertyName = "TcpPort" then
		portnum = "1433"
		tcpProperty.SetStringValue(portnum)
	end if
next
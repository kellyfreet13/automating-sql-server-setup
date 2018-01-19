# If users want to connect to the Site Audit database rfom a remote machine, 
# such as the case when installing multiple viewers in the same organization,
# you must:
#   set the startup type to automatic
#   start the service

Set-Service -Name SQLBrowser -StartupType Automatic
Restart-Service SQLBrowser
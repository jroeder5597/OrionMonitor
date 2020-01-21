Add-PSSnapin SwisSnapin

if($cred -eq $null){
$cred = Get-Credential #UA
}
#Connection to Orion database and initial SQL queries
$swis = Connect-Swis -Hostname "SWQL_DB_IP" -Credential $cred

$Orion = Get-SwisData -SwisConnection $swis -Query 'SELECT Nodes.Displayname, MIN(Nodes.Status) AS Status, MIN(Nodes.StatusDescription) AS Description, MIN(Nodes.NodeId) AS NodeID, MIN(NodeNotes.Note) AS Note, MIN(NodesCustomProperties.ResponsibleTeam) AS Team, MIN(NodeNotes.Timestamp) AS Time
FROM Orion.Nodes
LEFT Join Orion.NodesCustomProperties
ON Nodes.NodeID = NodesCustomProperties.NodeID 
LEFT JOIN Orion.NodeNotes
ON Nodes.NodeID = NodeNotes.NodeID
WHERE Status = "2" OR Status = "9"
GROUP BY Nodes.DisplayName'
#Query for unmanaged nodes, with filtering
$OrionUnmanage = Get-SwisData -SwisConnection $swis -Query "SELECT Nodes.Displayname
FROM Orion.Nodes
WHERE Status = '9' AND Nodes.DisplayName NOT LIKE '%MCM%' AND Nodes.DisplayName NOT LIKE '%FNGas%' AND Nodes.DisplayName NOT LIKE '%ION%'"
#Pull and filter audit messages based on unmanaged nodes
$orion2 = @()
$UserArray = @()
foreach($thing in $OrionUnmanage){
$orion2 += Get-SwisData -SwisConnection $swis -Query "SELECT TOP 1 AuditEventMessage, TimeLoggedUtc 
FROM Orion.AuditingEvents WHERE AuditEventMessage LIKE '%$thing%' AND AuditEventMessage like '%unmanage%' AND AuditEventMessage NOT LIKE '%ua57172%' ORDER BY TimeLoggedUtc DESC"
}
#Create hashmap with value user account and keys of unmanaged servers based on audit message
$UserHash = @{}
foreach($message in $Orion2){
$UserObject = New-Object -TypeName PSObject
    if($message.AuditEventMessage -like "*SECZONE*"){
        $User = $message.AuditEventMessage.Split(" ") | Select-String -Pattern "UA"
        $User = $User -split('SECZONE\\')
        $Server = $message.AuditEventMessage.Split() | Select-String -Pattern "TUS","HQD","EMS","KGM","SGS","ENG"
        #index of 1 to grab user value from split array
        $UserObject | Add-Member -MemberType NoteProperty -Name "Name" -Value $User[1] -Force
        $UserObject | Add-Member -MemberType NoteProperty -Name "Server" -Value $Server -Force
        #$UserHash.add($UserObject.Name, $UserObject.Server)
        if(!($UserHash.ContainsKey($User[1]))){
            $key = $user[1]
            $value = $Server
            $UserHash.Add($key, $value)
        }else{
            $UserHash[$user[1]] = $UserHash[$user[1]] + ", " + $Server
        }
        $UserArray += $UserObject
    }   
}
$key = $UserHash.Keys
#Create email for each user with associated unmanaged nodes
foreach($key in $UserHash.Keys){
$Mail = Get-ADUser $key -Properties * | Select mail | Out-String
$body = @()
$body += $key + " has unmanaged the following nodes, please provide a status: " + "<br><br>"
$body += $UserHash.$key
$body = $body | Out-String

$mailPeople=@{
    To = $Mail
    From = "jroeder@tep.com"
    CC = "jroeder@tep.com"
    Subject = "Reply Requested: Unmanaged Orion Nodes"
    SMTPServer = 'smtpmail.unisource.corp'
    Body = $body
    BodyAsHTML = $True
    Credential = $cred
    }
    Write-Host $Mail
    Write-Host $UserHash.$key
    Send-MailMessage @mailPeople
}
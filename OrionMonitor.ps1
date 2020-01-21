Add-PSSnapin SwisSnapin

if($cred -eq $null){

$cred = Get-Credential #UA

}

$swis = Connect-Swis -Hostname "10.100.10.86" -Credential $cred

$Orion = Get-SwisData -SwisConnection $swis -Query 'SELECT Nodes.Displayname, MIN(Nodes.Status) AS Status, MIN(Nodes.StatusDescription) AS Description, MIN(Nodes.NodeId) AS NodeID, MIN(NodeNotes.Note) AS Note, MIN(NodesCustomProperties.ResponsibleTeam) AS Team, MIN(NodeNotes.Timestamp) AS Time

FROM Orion.Nodes

LEFT Join Orion.NodesCustomProperties

ON Nodes.NodeID = NodesCustomProperties.NodeID 

LEFT JOIN Orion.NodeNotes

ON Nodes.NodeID = NodeNotes.NodeID

WHERE Status = "2" OR Status = "9"

GROUP BY Nodes.DisplayName'

$System = @()

$NATS = @()

foreach($node in $Orion){

    if($node.Team -eq "System"){

    $System += $node

    }else{

    }

    if($node.Team -eq "NATS" -and $node.Team -ne "System" -and $node.Displayname -notlike "*MCM*" -and $node.Displayname -notlike "*FNGas*" -and $node.Displayname -notlike "*GT1*"){

        $NATS += $node

        }else{

        }

}

$date = (Get-Date -Format -MM-dd-yyyy).ToString()

$System + $NATS | Out-File ("G:\IS_Infrastructure\Shared\InfrastructureOps\Work Folders\Jake\Orion Nodes\Nodes" + $date + ".csv")

$System = $System | ConvertTo-Html

$NATS = $NATS | ConvertTo-Html

$bodySystem = 'Please provide a status for the following nodes: ' + $System

$mailSystems=@{

To = "jroeder@tep.com"

From = "jroeder@tep.com"

CC = "jroeder@tep.com"

Subject = "Reply Requested: Systems Orion Nodes"

SMTPServer = 'smtpmail.unisource.corp'

Body = $bodySystem

BodyAsHTML = $True

Credential = $cred

}

Send-MailMessage @mailSystems

$bodyNATS = 'Please provide a status for the following nodes: ' + $NATS

 $mailNATS=@{

To = "jroeder@tep.com"

From = "jroeder@tep.com"

CC = "jroeder@tep.com"

Subject = "Reply Requested: NATS Orion Nodes"

SMTPServer = 'smtpmail.unisource.corp'

Body = $bodyNATS

BodyAsHTML = $True

Credential = $cred

}

Send-MailMessage @mailNATS
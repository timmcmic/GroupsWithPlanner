#Declare variables utilized through script.

$functionCommandTest = @()
$graphGroupsPowershellModule = "Microsoft.Graph.Groups"
$graphPlannerPowershellModule = "Microsoft.Graph.Planner"
$graphScopesHighestPrivilage = @("Group.Read.All", "Group.ReadWrite.All", "Tasks.ReadWrite")
$unifiedGroups = @()
$unifiedGroupsObjects = @()

#Ensure that the appropriate graph modules are present.

#Test for Microsoft.Graph.Groups

write-host "Testing Microsoft.Graph.Groups command presence..."

$functionCommandTest = get-command -module $graphGroupsPowershellModule

write-host $functionCommandTest.count

if ($functionCommandTest.count -eq 0)
{
    write-error "Microsoft graph groups powershell module not found.  Run install-module Microsoft.Graph.Groups"
}
else 
{
    write-host "Microsoft Graph Groups Powershell module found and commands are loaded."
}

$functionCommandTest = @()

write-host "Testing Microsoft.Graph.Planner command presence..."

$functionCommandTest = get-command -module $graphPlannerPowershellModule

write-host $functionCommandTest.count

if ($functionCommandTest.count -eq 0)
{
    write-error "Microsoft graph planner powershell module not found.  Run install-module Microsoft.Graph.Planner"
}
else 
{
    write-host "Microsoft Graph Planner Powershell module found and commands are loaded."
}

write-host "Establishing graph connection with required scopes."

try {
    connect-MGGraph -scopes $graphScopesHighestPrivilage -errorAction Stop

    write-host "Graph connection with required scopes successful."
}
catch {
    write-error "Graph connection with required scopes failed."
}

#Gather all groups in entraID with the tag unified.  

write-host "Gather all unified groups present in Entra ID."

$unifiedGroups = Get-MgGroup -All -expandProperty Owners | where {$_.groupTypes -contains "Unified"}

if ($unifiedGroups.count -gt 0)
{
    write-host "Unified groups were found for processing."
    write-host $unifiedGroups.count
}
else 
{
    write-error "No unified groups were found for processing.  Unified groups are required in order to utilize this code."
}

#Iterate through each of the groups and create a small powershell object that will presented information on the group.

foreach ($unifiedGroup in $unifiedGroups)
{
    #Determine the owners.

    write-host ("Procesing group object id: "+$unifiedGroup.id)

    if ($unifiedGroup.Owners.count -ge 1)
    {
        $owners = $unifiedGroup.Owners.AdditionalProperties.userPrincipalName -join ","
        $ownersCount = $unifiedGroup.Owners.count
    }
    else 
    {
        $owners = "None"
        $ownersCount = 0
    }

    $members = Get-MgGroupMember -GroupId $unifiedGroup.ID

    $hasPlanner = "Error"

    write-host "Determine if the group has any plans assigned to it..."

    try {
        $plans = Get-MGGroupPlannerPlan -groupID $unifiedGroup.id -errorAction Stop

        write-host "Able to execute planner graph command."

        if ($plans.count -gt 0)
        {
            write-host "Group has one or more plans." -ForegroundColor Green
            $hasPlanner = "Yes"
        }
        else 
        {
            write-host "Group does not have any plans." -ForegroundColor Yellow
            $hasPlanner = "No"
        }
    }
    catch {
       write-host "Error obtaining planner information - this could be by design or access denied etc - usually means no plan present..." -ForegroundColor Red
    }

    $group = [PSCustomObject] @{
        GroupId         = $unifiedGroup.Id
        GroupName       = $unifiedGroup.DisplayName
        CreatedDateTime = $unifiedGroup.CreatedDateTime
        Mail            = $unifiedGroup.Mail
        Visibility      = $unifiedGroup.Visibility
        Owners          = $owners
        OwnersCount     = $owners.count
        MembersCount    = $members.count
        HasPlanner      = $hasPlanner 
     }

     $unifiedGroupsObjects += $group
}

$unifiedGroupsObjects | Export-Csv .\groupsPlannerInformation.csv
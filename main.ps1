#region variables
$root = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition);

$NO_ACCESS_GROUP_PREFIX = "Sample_Server_NoAccess";
$NO_ACCESS_GROUPS = @(@("Sample_Server_NoAccess_0","3AAE"),@("Sample_Server_NoAccess_1","3482"),@("Sample_Server_NoAccess_2","5CBE"),@("Sample_Server_NoAccess_3","69DA"));
$SERVER = “Sample_Server”;

$ORGANIZATIONS = @("Sample_Org_0","Sample_Org_1","Sample_Org_2","Sample_Org_3","Sample_Org_4","Sample_Org_5");
$PASSWORDS = @("sample_pwd_0","sample_pwd_1","sample_pwd_2","sample_pwd_3","sample_pwd_4","sample_pwd_5","sample_pwd_6");
$PASSWORD_USER_RUNNING_SCRIPT = "password";

$CERTIFIER_FILES_ROOT_PATH = "\\path\to\sample\certifier_files";
$CERTIFIER_FILES = @("Sample_Org_0\cert.id","Sample_Org_1\cert.id","Sample_Org_2\cert.id","Sample_Org_3\cert.id","Sample_Org_4\cert.id","Sample_Org_5\cert.id");

$prefered_nag = 3;

$history = New-Object System.Collections.ArrayList;
$date = Get-Date -Format "yyMMddHHmm"
$log = ($root + "\logs\log " + $date + ".txt");

#Lists of objects that hold information about a user. 
$active_users = New-Object System.Collections.ArrayList;
$none_active_users = New-Object System.Collections.ArrayList;
$undetermined_users = New-Object System.Collections.ArrayList;

#Adds a first entry to each list which is later used as a header when exporting a csv-file.
$active_users.Add("FullLotusName;Name;Username;Company;Expiration Date")>$null;
$none_active_users.Add("FullLotusName;Name;Username;Company;Expiration Date")>$null;
$undetermined_users.Add("FullLotusName;Name;Username;Company;Expiration Date")>$null;


$error_codes = @(
"Felkod(0) – Powershell not running in 32-bit.",
"Felkod(1) – Active Directory module not found."
"Felkod(2) – No Lotus Notes Client installed."
"Felkod(3) – Could not open ID-file."

)


#endregion
#region methods

#[UserIsDisabled]
#A method that checks if a user is a member of a NoAccess group.
function UserIsDisabled($user)
{
    foreach($group in $user.Groups)
    {
        if(($group.name -match $NO_ACCESS_GROUP_PREFIX))
        {
            return $true;
        }
    }
    return $false;
}
#[RemoveuserFromGroup]
#Fetches the list of members from a specific group.
#Creates a copy of the list with a specific user omitted.
#Replaces the original list with the new list and saves the document.
function RemoveuserFromGroup($session, $target_user, $target_group)
{  
    $document = $db.GetDocumentByID($target_group.ID);
   
    $values = $document.GetItemValue("Members");
    
    $members_original = $values;

    [Int32]$len = $members_original.Count-1;

    if($len -lt 1)
    {
        return;
    }

    $members_updated = [System.Array]::CreateInstance([string],$len); 

    [int32]$index = 0;
    foreach($member in $members_original)
    {                    
        if($member -ne $target_user)
        {
            $members_updated[$index] = $member;
            $index++;
        }                    
    }

    $document.ReplaceItemValue("Members", $members_updated);

    $temp = $document.Save($true,$true);
}
#[AddToHistory]
#Takes in a string and writes it to the log file.
#Writes the string to the screen if specified not to.
function AddToHistory($line, $omit_output)
{
    $history.Add($line)>$null;
    Add-Content -Path $log -Value $line

    if(!$omit_output)
    {        
        write-host $line;
    }

}
#[GetExpiringUsers]
#Fetch all entries in the view "People (other)\Certificate Expiration" from the database names.nsf.
#Loop all items in the docoment from the current entry.
#If the Name propetry on the item is "FullName", fetch the items Text property.
#Use the Evaluate function to resolve the expiration date
#Create an object and put it into a list.
#Save list to csv-file.
function GetExpiringUsers($session, $max_date)
{    
    AddToHistory -line ("Fetching users expiring in " + $max_date + "...") -omit_output $false;

    $user_list = New-Object System.Collections.ArrayList;

    $user_list.Add("FullLotusName;Name;Expiration Date")>$null;

    $NotesAdressbook = $session.GetDatabase($SERVER, "names.nsf", 0);

    $views = $NotesAdressbook.Views();

    $index = 0;

    foreach($v in $views)
    {         
        $entries = $v.AllEntries();
    
        if($v.name -eq "People (other)\Certificate Expiration")
        { 
            $entries = $v.AllEntries()
        
            $e = $entries.GetFirstEntry();
        
            while(($e -ne $null))
            {           
                $doc = $e.Document();

                $test = $doc.Items();

                foreach($t in $test)
                {
                    if($t.Name -eq "FullName")
                    {
                        $s = $t.Text.Split('/O');
                        $name = $s[0].Substring(3,$s[0].Length-3);                    
                    
                        $s = $t.Text.Split(';');
                        $full_name = $s[0];     
                    } 
                
                }
                $date = [string]$session.Evaluate('@Date(@Certificate([Expiration];Certificate))', $doc);

                $date = $date.Replace(" ","");

                $mdy = $date.Split('/');

                $date = ($mdy[2] + "-" + $mdy[0] + "-" + $mdy[1]);
            
                $date = $date.Replace("00:00:00","");

                if($date -lt $max_date)
                {
                    $line = ($full_name + ";" + $name + ";" + $date);
                    $user_list.Add($line)>$null;  
                                  
                }                

                $e = $entries.GetNextEntry($e);
            }

            break;
        }
   

    }
    [Int32]$result = $user_list.Count-1;
    AddToHistory -line ($result.ToString()  + " users found") -omit_output $false;

    $path = ($root + "/batch/expiring_users.csv");

    if(![System.IO.File]::Exists($path))
    {
        New-Item $path;
    }

    Set-Content -Value $user_list -Path ($root + "/batch/expiring_users.csv") -Encoding UTF8;
}
#[EvaluateExpiredUsers]
#Fetch list of expiring users
#Loop expiring users and search for a match in Active Directory
#Check if user has allready been deleted
#Add user to 
function EvaluateExpiredUsers($deleted_users)
{
    $list = Get-Content -Path ($root + "/batch/expiring_users.csv") -Encoding UTF8 | ConvertFrom-Csv -Delimiter ';';    

    [int32]$total = $list.Length;

    [int32]$n = 1;

    AddToHistory -line "C = Awaiting Certification    R = Awaiting Removal    I = Ignore" -omit_output $false; 
    write-host ""

    foreach($entry in $list)
    {
        $name = $entry.FullLotusName.Split('/')[0].Replace('CN=','');
        try
        {
            $user = Get-ADUser -Filter {DisplayName -like $name} -Properties DisplayName,SamAccountName, enabled, Company | select DisplayName,SamAccountName, enabled, Company;
        }
        catch
        {
            AddToHistory -line ("Not found in AD " + $name) -omit_output $true;  
        }
    
        if(($user.SamAccountName -eq $null) -or ($user.SamAccountName -eq ""))
        {
            $username = "unknown";
        }
        else
        {
            $username = $user.SamAccountName;
        }

        if(($user.Company -eq $null) -or ($user.Company -eq ""))
        {
            $company = "unknown";
        }
        else
        {
            $company = $user.Company;
        }

        if($user.enabled -eq $true)
        {
            $str = ($entry.FullLotusName + ";" + $entry.Name + ";" + $username + ";" + $company + ";" + $entry.'Expiration Date');
            $active_users.Add($str)>$null;            
            AddToHistory -line ("[" + $n + "/" +,$total + "][C] " + $entry.FullLotusName) -omit_output $false;   
        }
        else
        {
            if($username -eq "unknown")
            {
                $allready_removed = $false;
                foreach($user in $deleted_users)
                {
                    if($user -eq $entry.FullLotusName)
                    {
                        $allready_removed = $true;
                        break;
                    }
                }

                if(!$allready_removed)
                {
                    $str = ($entry.FullLotusName + ";" + $entry.Name + ";" + $username + ";" + $company + ";" + $entry.'Expiration Date');
                    $undetermined_users.Add($str)>$null;                 
                    AddToHistory -line ("[" + $n + "/" +,$total + "][I] " + $entry.FullLotusName) -omit_output $false; 
                }                             
            }
            else
            {
                $allready_removed = $false;
                foreach($user in $deleted_users)
                {
                    if($user -eq $entry.FullLotusName)
                    {
                        $allready_removed = $true;
                        break;
                    }
                }

                if(!$allready_removed)
                {
                    $str = ($entry.FullLotusName + ";" + $entry.Name + ";" + $username + ";" + $company + ";" + $entry.'Expiration Date');
                    $none_active_users.Add($str)>$null;                
                    AddToHistory -line ("[" + $n + "/" +,$total + "][R] " + $entry.FullLotusName) -omit_output $false;   
                }
            }
        }      
        
        $n++;
    }

    $path_active_users = ($root + "/batch/certification.csv");
    $path_none_active_users = ($root + "/batch/removal.csv");
    $path_undetermined_users = ($root + "/batch/unresolved.csv");

    if(![System.IO.File]::Exists($path_active_users))
    {
        New-Item $path_active_users;
    }
    if(![System.IO.File]::Exists($path_none_active_users))
    {
        New-Item $path_none_active_users;
    }

    Set-Content -Value $active_users -Path $path_active_users -Encoding UTF8;
    Set-Content -Value $none_active_users -Path $path_none_active_users -Encoding UTF8;
    Set-Content -Value $undetermined_users -Path $path_undetermined_users -Encoding UTF8;


    AddToHistory -line ("") -omit_output $false;
    AddToHistory -line ("[Cross Referencing RESOLVED] Found " + ($active_users.Count-1) + " users awaiting certification") -omit_output $false;
    AddToHistory -line ("[Cross Referencing RESOLVED] Found " + ($none_active_users.Count-1) + " users awaiting removal") -omit_output $false;
    AddToHistory -line ("[Cross Referencing UNRESOLVED] Found " + ($undetermined_users.Count-1) + " users with no equivalence in Active Directory") -omit_output $false;

}

function CertifyUser($session, $full_lotus_name, $progress)
{
    $date = (get-date).AddYears(2);

    $nap = $session.CreateAdministrationProcess($SERVER);

    $nap.CertificateExpiration = $session.CreateDateTime($date);
    
    $organization = $full_lotus_name.Split('/')[1].Replace('OU=','');    

    $index = 0;
    $pwd = "bergendahls";

    for($i = 0;$i -lt $ORGANIZATIONS.length;$i++)
    {
        if($organization -eq $ORGANIZATIONS[$i])
        {
            $index = $i;
            $pwd = $PASSWORDS[$i];
        }
    }   


    if($index -ge 0 -and $index -lt 6)
    {   
        $nap.CertifierFile = $CERTIFIER_FILES_ROOT_PATH + $CERTIFIER_FILES[$index];

        $nap.CertifierPassword = $pwd;
        $flag = $true;

        try
        {
            $c = $nap.RecertifyUser($full_lotus_name);
        }
        catch
        {
            $flag = $false;
        }
        
        if($flag)
        {
            AddToHistory -line ("[Certification Success] User:" + $full_lotus_name + " (" + $CERTIFIER_FILES[$index] + ")") -omit_output $false;
        }
        else
        {
            AddToHistory -line ("[Certification FAILED] User:" + $full_lotus_name + " (" + $CERTIFIER_FILES[$index] + ")") -omit_output $false;
        }

        
    }
    else
    {
        AddToHistory -line ("[Certification FAILED] User:" + $full_lotus_name ) -omit_output $false;
    }
    
}
function BreakSegment($title)
{
    AddToHistory -line (" ") -omit_output $false;
    AddToHistory -line ("================================ " + $title + " ================================") -omit_output $false;
    AddToHistory -line (" ") -omit_output $false;
}
function CertifyUsersInList($list)
{
    [Int32]$n = 1;
    [Int32]$total = $list.Count;

    foreach($entry in $list)
    {     
        $progress = "" + $n + "/" + $total + "";
        $n++;
        CertifyUser -session $session -full_lotus_name $entry.FullLotusName -progress $progress;
    }

}
function RemoveUsersInList($list)
{
    foreach($user in $list)
    {
        if(!(UserIsDisabled -user $user))
        {
            foreach($group in $no_access_groups)
            {
                if($group.Name -eq $NO_ACCESS_GROUPS[$prefered_nag][0])
                {
               
                    $members = [System.Array]::CreateInstance([string],1); 

                    $refactored_username = $user.FullLotusName;

                    $refactored_username = $refactored_username.Replace('CN=','');

                    $refactored_username = $refactored_username.Replace('OU=','');

                    $refactored_username = $refactored_username.Replace('O=','');
                
                    $members[0] = $refactored_username;
                    $nap.AddGroupMembers($group.Name, $members)>$null;
                                
                    AddToHistory -line ("Added [" + $refactored_username + "] to group [" + $group.Name + "]") -omit_output $false;
                    $group.MemberCount++;               
                }
            } 
        }   
        foreach($group in $user.Groups)
        {      
            if(!($group.Name -match $NO_ACCESS_GROUP_PREFIX))
            {
                 AddToHistory -line ("Removed [" + $user.Username + "] from group [" + $group.Name + "]") -omit_output $false;
             
                 RemoveuserFromGroup -session $session -target_user $user.FullLotusName -target_group $group;
            }  
        }
    }
}

#endregion

#ENTRY POINT

if((Get-Module -ListAvailable -Name ActiveDirectory) -eq $null)
{
    AddToHistory -line $error_codes[1] -omit_output $false;
    read-host "Press Enter to exit."
    return;
}
if([IntPtr]::size -ne 4)
{
    AddToHistory -line $error_codes[0] -omit_output $false;
    read-host "Press Enter to exit."
    return;
}


clear;

Clear-Content -Path ($root + "/batch/expiring_users.csv")
Clear-Content -Path ($root + "/batch/removal.csv")
Clear-Content -Path ($root + "/batch/certification.csv")

New-Item -Path $log;

try
{
    $session = New-Object -ComObject Lotus.NotesSession -ErrorAction Stop
}
catch
{
    AddToHistory -line $error_codes[2] -omit_output $false;
    read-host "Press Enter to exit."
    return;
}
try
{

    $session.Initialize($PASSWORD_USER_RUNNING_SCRIPT)
}
catch
{
    AddToHistory -line $error_codes[3] -omit_output $false;
    read-host "Press Enter to exit."
    return;
}

$db = $session.GetDatabase($SERVER, "names.nsf", 0);


$deleted_users = New-Object System.Collections.ArrayList;

foreach($group in $NO_ACCESS_GROUPS)
{
    $document = $db.GetDocumentByID($group[1]);

    $members = $document.GetItemValue("Members");

    foreach($member in $members)
    {
        $deleted_users.Add($member)>$null;
    }

}

BreakSegment -title "Compiling users";


$date_max = (get-date -Day 30).AddMonths(1);

$date_max = $date_max.tostring(“yyyy-MM-dd”);

GetExpiringUsers -session $session -max_date $date_max;

BreakSegment -title "Cross Referencing";

EvaluateExpiredUsers -deleted_users $deleted_users;

write-host "";

$confirm = read-host "Confirm removal [Y/N]";

$list_removal = New-Object System.Collections.ArrayList;

if($confirm.ToLower() -eq "y")
{
    $list = Get-Content -Path ($root + "/batch/removal.csv") -Encoding UTF8 | ConvertFrom-Csv -Delimiter ';'


    foreach($entry in $list)
    {  
        $obj = New-Object -Type PSObject -Property @{
            'FullLotusName' = $entry.FullLotusName
            'Username' = $entry.Name
            'Company' = $entry.Company
            'Expiration Date' = $entry.'Expiration Date'     
            'Groups' = New-Object System.Collections.ArrayList
        }

        $list_removal.Add($obj)>$null;
    }


    $nap = $session.CreateAdministrationProcess($SERVER);

    $db = $session.GetDatabase($SERVER, "names.nsf", 0);

    BreakSegment -title "Disable Users";

    $groups = $db.GetView('($VIMGroups)');

    $entries = $groups.AllEntries()
            
    $group = $entries.GetFirstEntry();

    while(($group -ne $null))
    { 
        [string]$group_name = $group.Document.GetItemValue("ListName");
        [string]$group_id = $group.NoteID;
        $document = $db.GetDocumentByID($group.NoteID);
   
        $members = $document.GetItemValue("Members");


        foreach($member in $members)
        {
            foreach($entry in $list_removal)
            {
                if($entry.FullLotusName -eq $member)
                {
                    $obj = New-Object -Type PSObject -Property @{
                        'Name' = $group_name
                        'ID' = $group_id
                    }

                    $entry.Groups.Add($obj)>$null;
                }
            }
        }        

        $group = $entries.GetNextEntry($group); 
    }

    $no_access_groups = New-Object System.Collections.ArrayList

    for($i = 0; $i -lt $NO_ACCESS_GROUPS.Length; $i++)
    {
        $document = $db.GetDocumentByID($NO_ACCESS_GROUPS[$i][1]);
        $members = $document.GetItemValue("Members");

        $obj = New-Object -Type PSObject -Property @{'Name' = $NO_ACCESS_GROUPS[$i][0];'ID' = $NO_ACCESS_GROUPS[$i][1];'MemberCount' = [Int32]$members.Count}
        $no_access_groups.Add($obj)>$null;    
    }


    RemoveUsersInList -list $list_removal;
}


BreakSegment -title "Certification";


$confirm = read-host "Confirm certification [Y/N]";

$list_certification = New-Object System.Collections.ArrayList;

if($confirm.ToLower() -eq "y")
{
    $list = Get-Content -Path ($root + "/batch/certification.csv") -Encoding UTF8 | ConvertFrom-Csv -Delimiter ';'
    
    foreach($entry in $list)
    {  
        $obj = New-Object -Type PSObject -Property @{
            'FullLotusName' = $entry.FullLotusName
            'Username' = $entry.Name
            'Company' = $entry.Company
            'Expiration Date' = $entry.'Expiration Date'     
            'Groups' = New-Object System.Collections.ArrayList
        }

        $list_certification.Add($obj)>$null;
    }

    CertifyUsersInList -list $list_certification
}

BreakSegment -title "Override";

$final_sweep = read-host "Manually override unresolved users [Y/N]";

$list_certification.Clear();

$list_removal.Clear();

if($final_sweep.ToLower() -eq "y")
{
    $list = Get-Content -Path ($root + "/batch/unresolved.csv") -Encoding UTF8 | ConvertFrom-Csv -Delimiter ';'
    $list_undetermined = New-Object System.Collections.ArrayList
    foreach($entry in $list)
    {  
        $obj = New-Object -Type PSObject -Property @{
            'FullLotusName' = $entry.FullLotusName
            'Username' = $entry.Name
            'Company' = $entry.Company
            'Expiration Date' = $entry.'Expiration Date'     
            'Groups' = New-Object System.Collections.ArrayList
        }
        $list_undetermined.Add($obj)>$null;
    }

    foreach($user in $list_undetermined)
    {
        $q = read-host ($user.FullLotusName  + "   | 0 = Ignore | 1 = Disable | 2 = Recertify |     ")

        $n = 0;

        if([Int32]::TryParse($q, [ref]$n))
        {
            if($n -gt -1 -and $n -lt 3)
            {
                if($n -eq 2)
                {
                     $obj = New-Object -Type PSObject -Property @{
                        'FullLotusName' = $user.FullLotusName
                        'Username' = $user.Name
                        'Company' = $user.Company
                        'Expiration Date' = $user.'Expiration Date'     
                        'Groups' = New-Object System.Collections.ArrayList
                    }
                    $list_certification.Add($obj)>$null;
                }
                elseif($n -eq 1)
                {
                    $obj = New-Object -Type PSObject -Property @{
                        'FullLotusName' = $user.FullLotusName
                        'Username' = $user.Name
                        'Company' = $user.Company
                        'Expiration Date' = $entry.'Expiration Date'     
                        'Groups' = New-Object System.Collections.ArrayList
                    }
                    $list_removal.Add($obj)>$null;
                }
            }
        }
    }
    
    BreakSegment -title "Confirm override";

    foreach($user in $list_certification)
    {
        AddToHistory -line ($user.FullLotusName + " (" + $user.Username + ") will be recertified.") -omit_output $false;
    }

    foreach($user in $list_removal)
    {
        AddToHistory -line ($user.FullLotusName + " (" + $user.Username + ") will be removed.") -omit_output $false;
    }

    $confirm = read-host "Confirm [Y/N]";


    if($confirm.ToLower() -eq "y")
    {   
        CertifyUsersInList -list $list_certification;
        RemoveUsersInList -list $list_removal;
    }


}

write-host "";

read-host "All done, press Enter to exit."




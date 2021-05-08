########################################################################################################
# Windows 8 Setup Script

# Name: win_setup
  $versionNum = "v1.0"
# Date: 29 May 2016
# Created By: Daniel Bentz (daniel.bentz@us.af.mil) and Kyle Wilson (kyle.wilson.13@us.af.mil)
# Release Notes:
# 20160529:  Created to automate MIP setup on Win8 for network connectivity and firewall rules.
########################################################################################################


$menuSelectionIndex = 0

########################################################################################################
#function to create and handle Menus
########################################################################################################
#region Dynamic Menu Creation
#// Super, fantastic, over-engineered, dynamic menu creation function.
#//   $title: accepts either single string, or array of strings, to display title above the menu item, array prints on more than one line
#//   $size: width in characters that the menu will be when created
#//   $symbol: character to iterate and print to form the 'window' GUI borders for the menu
#//   $items: 1x3 dimensional array of objects to generate menu selection items, their functions, and hot-keys
#//       -Menu Item name string
#//       -Menu Item function to be executed when selected, use "Break" if nothing desired
#//       -Menu Item selection hot-key, press this key on the keyboard to select the item, double-tap or enter to execute
function MakeMenu($title, $size, $symbol, $items)
{
    #// Reset the selection index to 0 every time a new menu is created
    $menuSelectionIndex = 0
    do {
        clear-host
        #// Clear some space out to start generating menu 'borders'
        $menuString = "`n`n $symbol"
        #// Print dynamic width border of specified $symbol
        for($i = 0; $i -le $size; $i++) {
            $menuString += $symbol
        }
        $menuString += "`n $symbol"
        #// Print empty spaces for 'inner' border area above title
        for($i = 1; $i -le $size - 1; $i++) {
            $menuString += " "
        }
        $menuString += " $symbol`n $symbol"
        #// Check to see if the $title was passed with more than one desired line (as an array)
        If($title -is [array]) {
            #// If the $title IS an array, then iterate through printing and centering each line
            for($i = 0; $i -lt $title.count; $i++) {
                $tempLength = $title[$i].length
                $tempTitleI = ($size / 2)-($tempLength / 2)
                for($j = 1; $j -le $tempTitleI; $j++) {
                    $menuString += " "
                }
                $menuString += $title[$i]
                for($j = 1; $j -le $tempTitleI - (1 - ($size - $tempLength) % 2); $j++) {
                    $menuString += " "
                }
                $menuString += " $symbol`n $symbol"
            }
        }
        else {
            #// Otherwise, assume single line title string and print it out
            $tempLength = ($size / 2)-($title.length / 2)
            for($i = 1; $i -le $tempLength; $i++) {
                $menuString += " "
            }

            $menuString += $title
            for($i = 1; $i -le $tempLength - (1 - ($size - $title.length) % 2); $i++) {
                $menuString += " "
            }
            $menuString += " $symbol`n $symbol"
        }
        #// Add in more blank spaces under the title
        for($i = 1; $i -le $size - 1; $i++) {
            $menuString += " "
        }
        #// Create border under menu title before selection items
        $menuString += " $symbol`n $symbol"
        for($i = 0; $i -le $size; $i++) {
            $menuString += $symbol
        }
        #// Iterate through array of selection items for the menu
        $itemNum = -1
        $itemChars = New-Object System.Collections.ArrayList
        if($items.count -lt 2) {

        }
        #// Loop through all the menu items
        ForEach($item In $items) {
            $itemNum++
            #// Add input hot-key into an array-list for later input simplification
            [void]$itemChars.Add($item[2])
            #// Print selection item hot-key inside of brackets (aesthetics)
            $itemName = "[ " + $item[2] + " ] " + $item[0]

            $menuString += "`n $symbol"
            #// Put some space between menu items
            for($i = 1; $i -le $size - 1; $i++) {
                $menuString += " "
            }
            $menuString += " $symbol`n $symbol"

            for($i = 1; $i -le $size / 8; $i++) {
                $menuString += " "
            }
            #// The entire purpose of this boolean is to help add some buffer space to the highlighted "selection" area by replacing a space
            $stupidBoolean = $false
            #// If the item being constructed is the selection index, print it with color to present it as "highlighted"
            if ($itemNum -eq $menuSelectionIndex) {
                #// This has to be done with a new Write-Host line, so print everything we have so far, then print the "highlighted" parts
                Write-Host -NoNewLine $menuString
                Write-Host -NoNewLine -ForegroundColor Black -BackgroundColor White $itemName ""
                #// Clear out the menuString after the highlight to print normal color again
                $menuString = ""
                $stupidBoolean = $true
            }
            else {
                #// If not the selected item, just print/add the Title and continue
                $menuString += $itemName
            }
            #// Find and left justify the menu item titles, at 1/8th of the menu size/width
            $tempLength = $size - (($size / 8) + $itemName.length)
            for($i = 1; $i -le $tempLength; $i++) {
                if($stupidBoolean -eq $true) {
                    #// That stupid boolean that negates a space on the highlighted menu item
                    $stupidBoolean = $false
                }
                else {
                    $menuString += " "
                }
            }
            $menuString += " $symbol`n $symbol"

            for($i = 1; $i -le $size - 1; $i++) {
                $menuString += " "
            }
            $menuString += " $symbol"
        }
        $menuString += "`n $symbol"
        for($i = 1; $i -le $size - 1; $i++) {
            $menuString += " "
        }
        $menuString += " $symbol`n $symbol"
        for($i = 0; $i -le $size; $i++) {
            $menuString += $symbol
        }
        $menuString += "`n"
        #// Finally print the bulk of the menuString we've been creating gets printed
        Write-Host -NoNewLine $menuString

        $selection = ""

        $Global:HostConnectCheck = 0
        #// Print input prompt string
        Write-Host "`n Please make a Selection: "
        #// Start key input verification, this DOES NOT WORK inside of the PowerShell ISE
        $keyInput = $host.UI.RawUI.ReadKey("NoEcho, IncludeKeyDown")
        If($keyInput.VirtualKeyCode -eq 13) {
            #// Is it the Enter key? If so, execute the function declared through the menu creation items array
            $selection = $items[$menuSelectionIndex][0]
            $tempFunc = $items[$menuSelectionIndex][1]
            #Write-Host "Enter " $tempFunc
            #Sleep 1
            Invoke-Expression $tempFunc
        }
        ElseIf($keyInput.VirtualKeyCode -eq 38) {
            #// Is it the Up Arrow? If so, decrease the selection index
            #Write-Host "Up"
            if($menuSelectionIndex -gt 0) { $menuSelectionIndex-- }
        }
        ElseIf($keyInput.VirtualKeyCode -eq 40) {
            #// Is it the Down Arrow? If so, increase the selection index
            #Write-Host "Down"
            if($menuSelectionIndex -lt $items.Count-1) { $menuSelectionIndex++ }
        }
        ElseIf($keyInput.Character -in $itemChars) {
            #// Is it a menuItem hot-key?
            For($i = 0; $i -lt $itemChars.Count; $i++) {
                $tempFunc = ""
                if($itemChars[$i] -eq $keyInput.Character) {
                    #// If so, is it already selected? If so, double-selection will activate the hot-key
                    if($menuSelectionIndex -eq $i) {
                        $tempFunc = $items[$i][1]
                    }
                    Else {
                        #// If not already selected, then select it
                        $menuSelectionIndex = $i
                    }
                    break
                }
                else {
                }
            }
            #// If we have designated a function to be called, now invoke
            if($tempFunc -ne "") {
                $selection = $items[$menuSelectionIndex][0]
                Invoke-Expression $tempFunc #// break is only breaking the immediate loop, so call it outside of any loops
            }
        }
        Else {
            Write-Host -ForegroundColor red -BackgroundColor black "Invalid input!"
        }
        #Write-Host $tempFunc $items[-1][1]
        #sleep 1
    } until($tempFunc -eq $items[-1][1]) #// Loop menu until last defined item ("back/exit") is selected...
    return $selection
}

#// Quick function to pause until any key is pressed
function PauseMenu()
{
    $Global:HostConnectCheck = 0
    Write-Host "`n Press any key to continue..."
    #// Start key input verification, this DOES NOT WORK inside of the PowerShell ISE
    $keyInput = $host.UI.RawUI.ReadKey("NoEcho, IncludeKeyDown")
}

#// Takes an input of numbers (min of 1 max of 36) to generate additional hotkey values (0-Z)
function GenerateHotKey($num)
{
    if($num -lt 10)
    {
        return [String]$num
    }
    else
    {
        return [char]($num+55) #65 = A, 66 = B, etc
    }
}
#endregion

########################################################################################################
#function to configure Firewall Options
########################################################################################################
#region FireWall

#These are the Windows Firewall port options that do not support any port configuration
$protocolTypesNoPortOptions = @("Any")

#These are the Windows Firewall port options that do support port configuration
$protocolTypesWithPortOptions = @("TCP","UDP")

#This global object is used with the dynamic menus to store the information about the rules
[PSCustomObject]$theRuleObj = $null

#This array is used to identify all of the possible rule parameters for the firewall functions
#This is being treated similar to an enumerator but since they don't work like I would like them too
#I am going to use an array instead
$ruleParams = @("Rule Name", "Enabled", "Direction",
                "Profiles", "LocalIP", "RemoteIP",
                "Action", "LocalPort", "RemotePort",
                "Protocol"
               )


#Function that calls the custom configured Firewall menu
#that allows the section of different functions
function FirewallMenu($quickSetup)
{
    if ($quickSetup) {

        MakeMenu "Windows Firewall Options" "70" "+" @(("Enable Firewall","EnableFirewall; Break","1"),
                                                   ("Disable Firewall","DisableFirewall; Break","2"),
                                                   ("Skip...","Break","X"))
    }
    else {
        MakeMenu "Windows Firewall Options" "70" "+" @(("Enable Firewall","EnableFirewall","1"),
                                                   ("Disable Firewall","DisableFirewall","2"),
                                                   ("Add Rule","AddFirewallRule","3"),
                                                   ("Edit Rule","PrepEditFirewallRule","4"),
                                                   ("List Rules","ListFirewallRules","5"),
                                                   ("Delete Rule","DeleteFirewallRule","6"),
                                                   ("Exit To Main Menu","Break","X"))
    }
}

#// TODO: Function to Toggle (Enable/Disable) all three Firewalls (Public/Private/Domain)
function ToggleFirewall()
{
    Write-Host
    $firewallState = netsh advFirewall show currentProfile state
    Write-Host $firewallState
    if ($firewallState = "OFF") {

    }
}

#Enables all three firewalls (public, private, and domain)
function EnableFirewall()
{
    Write-Host
    #Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled True  #this requires cmdlets to be installed to work. netsh should be shipped with Windows
    #for compatibility reasons netsh is used instead of the cmdlet

    $return1 = netsh advfirewall set publicprofile state on
    $return2 = netsh advfirewall set domainprofile state on
    $return3 = netsh advfirewall set privateprofile state on
    if($return1 -eq "Ok." -and $return2 -eq "Ok." -and $return3 -eq "Ok.")
    {
        Write-Host -BackgroundColor Black -ForegroundColor Green "Firewall Enabled"
        Sleep 1 #small pause so output can be read
    }
    else
    {
        Write-Host -BackgroundColor Black -ForegroundColor Red "Error enabling the firewall"
        Write-Host -BackgroundColor Black -ForegroundColor Red "Output from attempt to turn on publicprofile:"
        Write-Host -BackgroundColor Black -ForegroundColor Red $return1
        Write-Host
        Write-Host -BackgroundColor Black -ForegroundColor Red "Output from attempt to turn on domainprofile:"
        Write-Host -BackgroundColor Black -ForegroundColor Red $return2
        Write-Host
        Write-Host -BackgroundColor Black -ForegroundColor Red "Output from attempt to turn on privateprofile:"
        Write-Host -BackgroundColor Black -ForegroundColor Red $return3
        Write-Host
        Read-Host "Press enter to continue"
    }
}

#Disables all three firewalls (public, private, and domain)
function DisableFirewall()
{
    Write-Host #the newline is for aesthetics
    #Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled False  #this requires cmdlets to be installed to work. netsh should be shipped with Windows
    #for compatibility reasons netsh is used instead of the cmdlet

    $return1 = netsh advfirewall set publicprofile state off
    $return2 = netsh advfirewall set domainprofile state off
    $return3 = netsh advfirewall set privateprofile state off
    if($return1 -eq "Ok." -and $return2 -eq "Ok." -and $return3 -eq "Ok.")
    {
        Write-Host -BackgroundColor Black -ForegroundColor Yellow  "Firewall Disabled"
        Sleep 1 #small pause so output can be read
    }
    else
    {
        Write-Host -BackgroundColor Black -ForegroundColor Red "Error disabling the firewall"
        Write-Host -BackgroundColor Black -ForegroundColor Red "Output from attempt to turn off publicprofile:"
        Write-Host -BackgroundColor Black -ForegroundColor Red $return1
        Write-Host
        Write-Host -BackgroundColor Black -ForegroundColor Red "Output from attempt to turn off domainprofile:"
        Write-Host -BackgroundColor Black -ForegroundColor Red $return2
        Write-Host
        Write-Host -BackgroundColor Black -ForegroundColor Red "Output from attempt to turn off privateprofile:"
        Write-Host -BackgroundColor Black -ForegroundColor Red $return3
        Write-Host
        Read-Host "Press enter to continue"
    }
}

#This prints all of the firewall rules and asks for the rule by name to edit
#It then prints the editable options for the firewall rule through the menu
#system. When changes are made, it updates the rule, updates the global variable
#and reprints the editable options in the menu system. There is data validation
#that does a decent job at preventing rules from happening but it is not perfect.
#If the rule editing fails it will display the error and pause
function EditFirewallRule([PSCustomObject]$localRuleObject)
{
    $theRuleObj = $localRuleObject #update the global object to be able to edit it throughout the script (and in this case through the menu system)
    $oldRuleName = $theRuleObj."Rule Name" #preserve the old name for the command

    Clear-Host
    Do
    {
        #if you don't clear these, then it will retain its old information = lots of issues
        $ruleParamsData = @()
        $changedName = $false

        #Go through all of the different parameters and add it to an array that will be used to create the dynamic menu
        for($i = 0; $i -lt $ruleParams.Count; $i++)
        {
            $ruleParamsData += , @("$($ruleParams[$i]) = $($theRuleObj."$($ruleParams[$i])")", "Break", (GenerateHotKey($i+1)))
        }
        #make sure to add the back/exit at the end (necessary for the makemenu function to work properly)
        $ruleParamsData += , @("Back", "Break", "X")

        #this clears the previous selection so it starts at the first index
        $menuSelectionIndex = 0

        #this calls the makemenu with the rules parameters and values as the option. Once something is selected it is returned as a string
        $selectedItem = MakeMenu "Select the option you wish to change" "70" "+" $ruleParamsData

        #checks to make sure there is an equals symbol to modify the information to be able to process it below
        if($selectedItem.Contains(" = "))
        {
            $selectedItem = $selectedItem.Substring(0, $selectedItem.IndexOf(" = ")) #retains the name of the parameter selected but removes the value
        }

        #resets the value
        $dataToChange = ""

        #the following chained if statements are used to go through each parameter and verify the input is good through data validation

        #Checks to make sure the rule name is unique and is not set to all (a reserved word).
        if($selectedItem -eq "Rule Name")
        {
            Do
            {
                $dataToChange = Read-Host "Enter new rule name (cannot be ""all"") [current=$($theRuleObj.$selectedItem)]"
                $dataToChange = $dataToChange.Trim()
                $changedName = $true
                $isDuplicateRule = isDuplicateRuleName $dataToChange
                if($isDuplicateRule -and $oldRuleName -ne $dataToChange)
                {
                    Write-Host -BackgroundColor Black -ForegroundColor Yellow "A rule already exists with that name, choose a different name"
                }
            } While($dataToChange -eq "all" -or $isDuplicateRule -and $oldRuleName -ne $dataToChange)
        }
        #Checks to make sure the direction is either in or out
        elseif($selectedItem -eq "Direction")
        {
            Do
            {
                $dataToChange = Read-Host "Enter new rule direction (in/out) [current=$($theRuleObj.$selectedItem)]"
            } While($dataToChange -ne "in" -and $dataToChange -ne "out" -and $dataToChange -ne "")
        }
        #Checks to make sure the enabled flag is set to yet or no
        elseif($selectedItem -eq "Enabled")
        {
            Do
            {
                $dataToChange = Read-Host "Do you want to enable this rule (yes/no) [current=$($theRuleObj.$selectedItem)]"
            } While($dataToChange -ne "no" -and $dataToChange -ne "yes" -and $dataToChange -ne "")
        }
        #Checks to make sure the IP setting doesn't have a space ***terrible data validation***
        elseif($selectedItem -eq "RemoteIP")
        {
            Do
            {
                $dataToChange = Read-Host "Enter new remote IP(s) with no spaces separated by a comma (e.g. Any,8.8.8.8,1.1.1.1-9.9.9.9,192.168.0.0/21) [current=$($theRuleObj.$selectedItem)]"
            } While($dataToChange.Contains(" ") -and $dataToChange -ne "")
        }
        #Checks to make sure the IP setting doesn't have a space ***terrible data validation***
        elseif($selectedItem -eq "LocalIP")
        {
            Do
            {
                $dataToChange = Read-Host "Enter new local IP(s) with no spaces separated by a comma (e.g. Any,8.8.8.8,1.1.1.1-9.9.9.9,192.168.0.0/21) [current=$($theRuleObj.$selectedItem)]"
            } While($dataToChange.Contains(" ") -and $dataToChange -ne "")
        }
        #Checks to make sure the profile setting doesn't have a space ***terrible data validation***
        elseif($selectedItem -eq "Profiles")
        {
            Do
            {
                $dataToChange = Read-Host "Enter new profile(s) with no spaces separated by a comma (Domain,Private,Public) [current=$($theRuleObj.$selectedItem)]"
            } While($dataToChange.Contains(" ") -and $dataToChange -ne "")
        }
        #Checks to make sure the protocol is in one of the protocol types defined in the arrays (with/with no port options)
        #also will not allow protocol changes if the port settings will cause an error when trying to save the changes
        elseif($selectedItem -eq "Protocol")
        {

            Do
            {
                $dataToChange = Read-Host "Enter new protocol type (Any, TPC, UDP) [current=$($theRuleObj.$selectedItem)]"
                if($dataToChange -ne "" -and $dataToChange -In $protocolTypesNoPortOptions -and (($theRuleObj.LocalPort -ne "Any" -and $theRuleObj.LocalPort) -or ($theRuleObj.RemotePort -ne "Any" -and $theRuleObj.RemotePort)))
                {
                   Write-Host -BackgroundColor Black -ForegroundColor Yellow "The protocol entered does not allow port filtering options"
                   Write-Host -BackgroundColor Black -ForegroundColor Yellow "Change the LocalPort and/or RemotePort configuration before changing the Protocol for the rule"
                   $protocolDataVal = $true
                }
                else
                {
                    $protocolDataVal = $false
                }
            } While($dataToChange -notin $protocolTypesNoPortOptions -and $dataToChange -notin $protocolTypesWithPortOptions -and $dataToChange -ne "" -or $protocolDataVal)
        }
        #Checks to make sure the IP setting doesn't have a space ***terrible data validation***
        #also will not allow port changes if the protocol settings will cause an error when trying to save the changes
        elseif($selectedItem -eq "LocalPort")
        {
            if($theRuleObj.Protocol -in $protocolTypesWithPortOptions)
            {
                Do
                {
                    $dataToChange = Read-Host "Enter new local port(s) with no spaces separated by a comma (e.g. 80,443,1-65534) [current=$($theRuleObj.$selectedItem)]"
                } While($dataToChange.Contains(" ") -and $dataToChange -ne "")
            }
            else
            {
                Write-Host -BackgroundColor Black -ForegroundColor Yellow "The current protocol ($($theRuleObj.Protocol)) does not allow custom port filtering in the Windows Firewall"
                Sleep 2
            }
        }
        #Checks to make sure the IP setting doesn't have a space ***terrible data validation***
        #also will not allow port changes if the protocol settings will cause an error when trying to save the changes
        elseif($selectedItem -eq "RemotePort")
        {

            if($theRuleObj.Protocol -in $protocolTypesWithPortOptions)
            {
                Do
                {
                    $dataToChange = Read-Host "Enter new remote port(s) with no spaces separated by a comma (e.g. 80,443,1-65534) [current=$($theRuleObj.$selectedItem)]"
                } While($dataToChange.Contains(" ") -and $dataToChange -ne "")
            }
            else
            {
                Write-Host -BackgroundColor Black -ForegroundColor Yellow "The current protocol ($($theRuleObj.Protocol)) does not allow custom port filtering in the Windows Firewall"
                Sleep 2
            }
        }
        ##Checks to make sure the action setting is either allow or block
        elseif($selectedItem -eq "Action")
        {
            Do
            {
                $dataToChange = Read-Host "Enter new action (block/allow) [current=$($theRuleObj.$selectedItem)]"
            } While($dataToChange -ne "block" -and $dataToChange -ne "allow" -and $dataToChange -ne "")
        }


        #This is where the data is updated/created
        if($dataToChange -ne "" -and $dataToChange -ne $theRuleObj.$selectedItem)
        {

            if(-not $theRuleObj.$selectedItem) #if a null value is found (or the member does not exist) then add it to the object
            {
                #the !!!DeleteItem!!!: is a unique string that will allow a good chance to substring it out if
                #there is an error. The error handling is at the end of this function
                $beforeChange = "!!!DeleteItem!!!:$selectedItem"
                $theRuleObj | Add-Member -MemberType NoteProperty -Name $selectedItem -Value $dataToChange -Force
            }
            else #otherwise just update the current value
            {
                #if is not adding a member than hold the old value in case there was an error
                $beforeChange = $theRuleObj.$selectedItem
                $theRuleObj.$selectedItem = $dataToChange
            }

            #Used to dynamically get the name of the command to update the information obtained
            #from the output of the netsh command. Not all parameters need to be listed statically,
            #only the ones that won't work with their outputted name.
            switch($selectedItem)
            {
                "Rule Name" {$item = "name"}
                "Direction" {$item = "dir"}
                "Enabled" {$item = "enable"}
                "Profiles" {$item = "profile"}
                Default {$item =  $selectedItem}
            }
            #runs the command to update the rule and returns the results
            $result = netsh advfirewall firewall set rule name=$oldRuleName new $item=$dataToChange

            #Error handling
            if ($result -eq "Ok.")
            {
                #if the name was changed, then update the oldRuleName variable so future edits will work
                if($changedName)
                {
                    $oldRuleName = $dataToChange
                }
                Write-Host -BackgroundColor Black -ForegroundColor Green "Rule edited successfully"
                Sleep 1
            }
            else #If an error occurred, reverse the changes made to the variables/objects that are displayed through the menu system
            {
                #This is the unique string used to substring out the name of the parameter
                if($beforeChange -like "!!!DeleteItem!!!:*")
                {
                    $theRuleObj.PSObject.Properties.Remove($beforeChange.Substring($beforeChange.IndexOf(":") + 1))
                }
                else
                {
                    $theRuleObj.$selectedItem = $beforeChange
                }

                #Display the error and the command that was used
                Write-Host -BackgroundColor Black -ForegroundColor Red "Error editing rule"
                Write-Host -BackgroundColor Black -ForegroundColor Red $result
                Write-Host
                Write-Host "The command that was used is:"
                Write-Host netsh advfirewall firewall set rule name=$oldRuleName new $item=$dataToChange
                PauseMenu
            }
        }
    } Until($selectedItem -eq "Back") #since each edit updates the rule, only back is needed (no save is necessary)
    $theRuleObj = $null #make sure the global object is set to nothing (reset) to prevent possible bugs
}

#Searches for a rule by name and returns the netsh output
#that can be processed later
function SearchFirewallRule([String]$name)
{
    return netsh advfirewall firewall show rule name=$name
}

#This function prepares the firewall rule to be edited. It also displays
#a notification of duplicate names being used for the rules (which is not
#supported using this script/netsh). If there are duplicates then it gives
#the user the option to delete all of the duplicates and tells them to
#create a new one
function PrepEditFirewallRule()
{
    [PSCustomObject]$rules = GetFirewallRules
    if($rules -ne $null)
    {
        Clear-Host
        PrintFirewallRules($rules)
        $ruleName = Read-Host "Type the rule name that you wish to edit"

        if($ruleName.Trim() -ne "")
        {
            $ruleText = SearchFirewallRule($ruleName) #Returns the Netsh output
            if($ruleText[1] -ne "No rules match the specified criteria.")
            {
                $ruleObj = GetPSObjectFromNetsh($ruleText)
                if($ruleObj.Count -gt 1) #duplicate rules section: note there has to be 2 or more for the .count to work
                {
                    Do
                    {
                        Clear-Host
                        Write-Host -BackgroundColor Black -ForegroundColor Yellow "WARNING"
                        Write-Host -BackgroundColor Black -ForegroundColor Yellow "There are" $ruleObj.Count "rules with the name $ruleName. All" $ruleObj.Count "duplicate rules will be deleted and you need to create a new one."
                        Sleep -Milliseconds 500
                        $input = Read-Host "Continue (y/n)?"
                    } Until($input -eq 'y' -or $input -eq 'n' -or $input -eq 'yes' -or $input -eq 'no')

                    if($input -eq 'y' -or $input -eq 'yes')
                    {
                        netsh advfirewall firewall delete rule name=$ruleName > $null
                    }
                }
                else
                {
                    EditFirewallRule($ruleObj) #This is where everything passed and the rule goes to the edit function
                }
            }
            else
            {
                Write-Host -ForegroundColor Yellow -BackgroundColor Black $ruleText[1] #prints there are no rules message
                Sleep 1
            }
        }
    }
    else
    {
        Write-Host
        Write-Host -ForegroundColor Yellow -BackgroundColor Black "***There are no rules***"
        Sleep 1
    }
}

#This prompts for a new rule name to create a new rule. It then assigns default values.
#Next it prints the editable options for the firewall rule through the menu
#system. When changes are made, it updates the local variables and reprints
#the editable options in the menu system. There is data validation
#that does a decent job at preventing rules from happening but it is not perfect.
#If the rule creation fails it will display the error and pause
function AddFirewallRule()
{
    Clear-Host

    #default values
    $direction = "in"
    $action = "block"
    $enabled = "yes"
    $remoteIP = "any"
    $localIP = "any"
    $profile = "domain,private,public"
    $protocol = "any"
    $localPort = ""
    $remotePort = ""

    #resets the duplicate rule flag
    $isDuplicateRule = $false


    #Checks to make sure the rule name is unique and is not set to all (a reserved word).
    Do
    {
        $ruleName = Read-Host "Enter the rule name (cannot be ""all"")"
        $ruleName = $ruleName.Trim()
        $isDuplicateRule = isDuplicateRuleName $ruleName
        if($isDuplicateRule)
        {
            Write-Host -BackgroundColor Black -ForegroundColor Yellow "A rule already exists with that name, choose a different name"
        }
    } While ($ruleName -eq "all" -or $isDuplicateRule)

    if($ruleName -ne "") #if the rule name is empty, cancel adding the rule (this skips the rest of the function)
    {
        Do
        {
            $ruleParamsData = @() #clear the array

            #this goes through each object in the global array (like an enum) and creates the dynamic menu based upon the switch
            #the ruleParamsData is the array that contains the arrays of inputs. Break exits the menu loop (we only care about
            #the return value). This returns the selection which is used later in the code. The generatehotkey function returns
            #1-9, A-Z for hotkey selections
            for($i = 0; $i -lt $ruleParams.Count; $i++)
            {
                Switch($ruleParams[$i])
                {
                    "Rule Name" { $ruleParamsData += , @("$($ruleParams[$i]) = $ruleName", "Break", (GenerateHotKey($i + 1))) }
                    "Enabled" { $ruleParamsData += , @("$($ruleParams[$i]) = $enabled", "Break", (GenerateHotKey($i + 1))) }
                    "Direction" { $ruleParamsData += , @("$($ruleParams[$i]) = $direction", "Break", (GenerateHotKey($i + 1))) }
                    "Profiles" { $ruleParamsData += , @("$($ruleParams[$i]) = $profile", "Break", (GenerateHotKey($i + 1))) }
                    "LocalIP" { $ruleParamsData += , @("$($ruleParams[$i]) = $localIP", "Break", (GenerateHotKey($i + 1))) }
                    "RemoteIP" { $ruleParamsData += , @("$($ruleParams[$i]) = $remoteIP", "Break", (GenerateHotKey($i + 1))) }
                    "Action" { $ruleParamsData += , @("$($ruleParams[$i]) = $action", "Break", (GenerateHotKey($i + 1))) }
                    "LocalPort" { $ruleParamsData += , @("$($ruleParams[$i]) = $localPort", "Break", (GenerateHotKey($i + 1))) }
                    "RemotePort" { $ruleParamsData += , @("$($ruleParams[$i]) = $remotePort", "Break", (GenerateHotKey($i + 1))) }
                    "Protocol" { $ruleParamsData += , @("$($ruleParams[$i]) = $protocol", "Break", (GenerateHotKey($i + 1))) }
                    Default { Write-Host "Missing switch argument in AddFirewallRule" }
                }
            }

            #These rules will always at the bottom of the list
            $ruleParamsData += , @("Save Rule", "Break", "S")
            $ruleParamsData += , @("Back", "Break", "X") #this is mandatory to be at the bottom for the script to function?

            #resets the menu selection so it will be at the top
            $menuSelectionIndex = 0
            #calls the function that makes the menu and returns the selected item
            $selectedItem = MakeMenu "Select the option you wish to change" "70" "+" $ruleParamsData

            #if the selected item contains an eqauls sign, substring the beginning part out (the rule parameter) to be used later
            if($selectedItem.Contains(" = "))
            {
                $selectedItem = $selectedItem.Substring(0, $selectedItem.IndexOf(" = "))
            }

            #clear out the dataToChange variable
            $dataToChange = ""


            #the following chained if statements are used to go through each parameter and verify the input is good through data validation

            #Checks to make sure the rule name is unique and is not set to all (a reserved word).
            if($selectedItem -eq "Rule Name")
            {
                Do
                {
                    $input = Read-Host "Enter new rule name (cannot be ""all"")"
                    $input = $input.Trim()
                    $isDuplicateRule = isDuplicateRuleName $input
                    if($isDuplicateRule)
                    {
                        Write-Host -BackgroundColor Black -ForegroundColor Yellow "A rule already exists with that name, choose a different name"
                    }
                } While($input -eq "all" -or $isDuplicateRule)
                if($input -ne "")
                {
                    $ruleName = $input
                }
            }
            #Checks to make sure the direction is either in or out
            elseif($selectedItem -eq "Direction")
            {
                Do
                {
                    $input = Read-Host "Enter new rule direction (in/out)"
                } While($input -ne "in" -and $input -ne "out" -and $input -ne "")
                if($input -ne "")
                {
                    $direction = $input
                }
            }
            #Checks to make sure the enabled flag is set to yet or no
            elseif($selectedItem -eq "Enabled")
            {
                Do
                {
                    $input = Read-Host "Do you want to enable this rule (yes/no)"
                } While($input -ne "no" -and $input -ne "yes" -and $input -ne "")
                if($input -ne "")
                {
                   $enabled = $input
                }
            }
            #Checks to make sure the IP setting doesn't have a space ***terrible data validation***
            elseif($selectedItem -eq "RemoteIP")
            {
                Do
                {
                    $input = Read-Host "Enter new remote IP(s) with no spaces separated by a comma (e.g. Any,8.8.8.8,1.1.1.1-9.9.9.9,192.168.0.0/21)"
                } While($input.Contains(" ") -and $input -ne "")
                if($input -ne "")
                {
                    $remoteIP = $input
                }
            }
            #Checks to make sure the IP setting doesn't have a space ***terrible data validation***
            elseif($selectedItem -eq "LocalIP")
            {
                Do
                {
                    $input = Read-Host "Enter new local IP(s) with no spaces separated by a comma (e.g. Any,8.8.8.8,1.1.1.1-9.9.9.9,192.168.0.0/21)"
                } While($input.Contains(" ") -and $input -ne "")
                if($input -ne "")
                {
                    $localIP = $input
                }
            }
            #Checks to make sure the profile setting doesn't have a space ***terrible data validation***
            elseif($selectedItem -eq "Profiles")
            {
                Do
                {
                    $input = Read-Host "Enter new profile(s) with no spaces separated by a comma (Domain,Private,Public)"
                } While($input.Contains(" ") -and $input -ne "")
                if($input -ne "")
                {
                    $profile = $input
                }
            }
            #Checks to make sure the protocol is in one of the protocol types defined in the arrays (with/with no port options)
            #also will not allow protocol changes if the port settings will cause an error when trying to save the changes
            elseif($selectedItem -eq "Protocol")
            {

                Do
                {
                    $input = Read-Host "Enter new protocol type (Any, TPC, UDP)"
                    if($input -ne "" -and $input -In $protocolTypesNoPortOptions -and (($localPort -ne "Any" -and $localPort -ne "") -or ($remotePort -ne "Any" -and $remotePort -ne "")))
                    {
                       Write-Host -BackgroundColor Black -ForegroundColor Yellow "The protocol entered does not allow port filtering options"
                       Write-Host -BackgroundColor Black -ForegroundColor Yellow "Change the LocalPort and/or RemotePort configuration before changing the Protocol for the rule"
                       $protocolDataVal = $true
                    }
                    else
                    {
                        $protocolDataVal = $false
                    }
                } While($input -notin $protocolTypesNoPortOptions -and $input -notin $protocolTypesWithPortOptions -and $input -ne "" -or $protocolDataVal)
                if($input -ne "")
                {
                    $protocol = $input
                }
            }
            #Checks to make sure the IP setting doesn't have a space ***terrible data validation***
            #also will not allow port changes if the protocol settings will cause an error when trying to save the changes
            elseif($selectedItem -eq "LocalPort")
            {
                if($protocol -in $protocolTypesWithPortOptions)
                {
                    Do
                    {
                        $input = Read-Host "Enter new local port(s) with no spaces separated by a comma (e.g. 80,443,1-65534)"
                    } While($input.Contains(" ") -and $input -ne "")
                    if($input -ne "")
                    {
                        $localPort = $input
                    }
                }
                else
                {
                    Write-Host -BackgroundColor Black -ForegroundColor Yellow "The current protocol ($protocol) does not allow custom port filtering in the Windows Firewall"
                    Sleep 2
                }
            }
            #Checks to make sure the IP setting doesn't have a space ***terrible data validation***
            #also will not allow port changes if the protocol settings will cause an error when trying to save the changes
            elseif($selectedItem -eq "RemotePort")
            {

                if($protocol -in $protocolTypesWithPortOptions)
                {
                    Do
                    {
                        $input = Read-Host "Enter new remote port(s) with no spaces separated by a comma (e.g. 80,443,1-65534)"
                    } While($input.Contains(" ") -and $input -ne "")
                    if($input -ne "")
                    {
                        $remotePort = $input
                    }
                }
                else
                {
                    Write-Host -BackgroundColor Black -ForegroundColor Yellow "The current protocol ($protocol) does not allow custom port filtering in the Windows Firewall"
                    Sleep 2
                }
            }
            #Checks to make sure the action setting is either allow or block
            elseif($selectedItem -eq "Action")
            {
                Do
                {
                    $input = Read-Host "Enter new action (block/allow)"
                } While($input -ne "block" -and $input -ne "allow" -and $input -ne "")
                if($input -ne "")
                {
                    $action = $input
                }
            }
        } Until($selectedItem -eq "Back" -or $selectedItem -eq "Save Rule")


        #if Save Rule was selected
        if($selectedItem -eq "Save Rule")
        {
            if($protocol -in $protocolTypesWithPortOptions) #if the protocol is in the portOptions array
            {
                $result = netsh advfirewall firewall add rule name="$ruleName" dir=$direction action=$action enable=$enabled remoteip=$remoteIP localip=$localIP profile=$profile protocol=$protocol localPort=$localPort remotePort=$remotePort
                $theCommand = "netsh advfirewall firewall add rule name=""$ruleName"" dir=$direction action=$action enable=$enabled remoteip=$remoteIP localip=$localIP profile=$profile protocol=$protocol localPort=$localPort remotePort=$remotePort"
            }
            else #if the protocol is in the noPortOptions array
            {
                $result = netsh advfirewall firewall add rule name="$ruleName" dir=$direction action=$action enable=$enabled remoteip=$remoteIP localip=$localIP profile=$profile protocol=$protocol
                $theCommand = "netsh advfirewall firewall add rule name=""$ruleName"" dir=$direction action=$action enable=$enabled remoteip=$remoteIP localip=$localIP profile=$profile protocol=$protocol"
            }
            if ($result -eq "Ok.") #if no errors
            {
                Write-Host -BackgroundColor Black -ForegroundColor Green "Rule created successfully"
                Sleep 1
            }
            else #if error(s) then do this
            {
                Write-Host -BackgroundColor Black -ForegroundColor Red "Error creating the rule"
                Write-Host -BackgroundColor Black -ForegroundColor Red $result
                Write-Host
                Write-Host "The command that was used is:"
                Write-Host $theCommand
                Read-Host "Press enter to continue"
            }
        }
    }
}

#used to determine if there are any rules that
#already exists with the same name. Since netsh
#does not really support modifying rules with the
#same name, I choose to not attempt to support it
#and will use this function to tell them to choose
#a new name
function isDuplicateRuleName([String]$ruleName)
{
    #the GetPSObjectFromNetsh function returns $null if there are no rules
    $ruleObj = GetPSObjectFromNetsh(SearchFirewallRule $ruleName)
    return ($ruleObj -ne $null) #if not null then there are duplicates, return true
}

#Displays the firewall rules and prompts for the rule name
#to be deleted.
function DeleteFirewallRule()
{
    $rules = GetFirewallRules
    if($rules -ne $null)
    {
        Clear-Host
        PrintFirewallRules($rules)
        $ruleName = Read-Host "Enter rule name to delete"
        Write-Host

        if(-not $ruleName.Trim() -eq "")
        {
            $return = netsh advfirewall firewall delete rule name=$ruleName

            if($return -eq "Ok.")
            {
                Write-Host -BackgroundColor Black -ForegroundColor Green "Rule Deleted"
                Sleep 1
            }
            else
            {
                Write-Host -BackgroundColor Black -ForegroundColor Red "Error deleting rule"
                Write-Host -BackgroundColor Black -ForegroundColor Red $return
                Write-Host
                Read-Host "Press enter to continue"
            }
        }
    }
    else
    {
        Write-Host
        Write-Host -ForegroundColor Yellow -BackgroundColor Black "***There are no rules***"
        Sleep 1
    }
}

#This function converts the returned output of netsh commands used in
#this script to a PSObject to be able to print the information
#in format-table. This function depends on the way the netsh function
#outputs information in Windows 8.1. The location of blank lines
#are used to determine the start of the next set of data and the
#colons are used to identify lines that have data
function GetPSObjectFromNetsh($netshData)
{
    $toReturn = @() #initializes the variable to null
    $tempObjectInfo = New-Object PSCustomObject
    for($i = 0; $i -lt $netshData.count; $i++) #loops through the array of text
    {
        if($netshData[$i] -eq "No rules match the specified criteria.")
        {
            return $null
        }
        else
        {
            #if there is a blank line and it is not the first line (line 0) then add the collection of information into the array
            if($netshData[$i].Trim() -eq "" -and $i -gt 0)
            {
                $toReturn += $tempObjectInfo #adds the PSCustom object to the toReturn array so it can be returned as a collection of rules instead of just 1 rule
                $tempObjectInfo = New-Object PSCustomObject #resets the object to a blank new one
            }
            #the only information that is needed is separated by a colon a lot of spaces. So only add that information
            if($netshData[$i].Contains(": "))
            {
                #the substring gets the name of the property from the left side of the colon and the value from the right side of the colon
                $tempObjectInfo | Add-Member -MemberType NoteProperty -Name $netshData[$i].Substring(0, $netshData[$i].IndexOf(": ")).Trim() -Value $netshData[$i].Substring($netshData[$i].IndexOf(": ") + (": ").Length).Trim()
            }
        }
    }
    return $toReturn
}

#Returns all of the firewall rules as a PSObject that can be printed with the format-table
function GetFirewallRules()
{
    #This captures the output of the command and stores it in the variable
    $rulesText = netsh advfirewall firewall show rule name=all
    return GetPSObjectFromNetsh($rulesText)
}

#This function requires the rules data from the GetFirewallRules function to print. This
#either displays the sorted rules or it prints the fact that there are no rules
function PrintFirewallRules([PSCustomObject]$rules)
{
    if($rules -ne $null)
    {
        #This sorts by the Rule Name object in the PSCustom object that was created
        #in the GetPSObjectFromNetsh function. The ExcludeProperty argument removes
        #the ones that you do not wish to display. This loop fixes unintended consequences
        #with having null values added (e.g. localports can be null and the sort-object
        #removes a null value if it is the first in the list)
        #$rules  | Format-Table -AutoSize
        foreach($rule in $rules)
        {
            #this loop goes through each member of the entire custom object and tests for null values in each rule
            $rules | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name | ForEach {
                if(-not $rule.$_)
                {
                    #if a null value is found (or the member does not exist) then add a blank string
                    $rule | Add-Member -MemberType NoteProperty -Name $_ -Value "" -Force
                }
            }
        }
        #this prints everything in a formatted table and it excludes properties that don't need to be displayed. Finally it sorts by the rule name
        $rules | Select-Object -Property * -ExcludeProperty Grouping, "Edge traversal" | Sort-Object -Property "Rule Name" | Format-Table -AutoSize
    }
    else
    {
        Write-Host -ForegroundColor Yellow -BackgroundColor Black "***There are no rules***"
    }
}

#Lists all of the firewall rules and pauses until a key is pressed
function ListFirewallRules()
{
    $rules = GetFirewallRules
    if($rules -ne $null)
    {
        Clear-Host
        PrintFirewallRules($rules)
        Read-Host "Press enter to continue"
    }
    else
    {
        Write-Host
        Write-Host -ForegroundColor Yellow -BackgroundColor Black "***There are no rules***"
        Sleep 1
    }

}

#endregion

########################################################################################################
#function to set IPAddress
########################################################################################################
#region IP Config
function IPMenu($selectedIF)
{
    if(-not $selectedIF) {
        $selectedIF = IFSelectionMenu("IP Address Configuration")
    }
    Write-Host "[$selectedIF]"
    if ($selectedIF -ne "" -and $selectedIF -ne "Back" -and $selectedIF -ne "Skip...") {
        #// Issue here is that DHCP can be enabled on the loopback interface which errors...
        Do {
            $useDHCP = Read-Host " Would you like to use DHCP? (Y | N)"
        } Until ($useDHCP -match "y" -or $useDHCP -match "n")
        if ($useDHCP -eq "y") {
            Try {
                $ErrorActionPreference = "Stop"
                Set-NetIPInterface $selectedIF -AddressFamily IPv4 -Dhcp Enabled -ErrorAction "Stop"
                Write-Host -ForegroundColor Green -BackgroundColor Black "`nYou have enabled DHCP on interface [$selectedIF]."
            }
            Catch {
                Write-Host -ForegroundColor Red -BackgroundColor Black "`nDHCP is already enabled ON [$selectedIF]."
            }
            Finally {
                $ErrorActionPreference = "Continue"
            }
        }
        Elseif ($useDHCP -eq "n") {
            #// Do some validation of the entered IP addresses, see if it satisfies IPAddress AND count the octects since the ipaddress doesn't...
            Do {
                #// Validate IP address
                $isValid = $false
                $IP = Read-Host " Enter IP address: "
                $isValid = $ip -as [ipaddress] -and $IP.split(".").Count -eq 4
            } until ($isValid -eq $true) #Until ($IP -as [ipaddress])
            Do {
                #// Validate subnet mask
                $isValid = $false
                $NETMASK = Read-Host " Enter your subnet mask: "
                $isValid = $NETMASK -as [ipaddress] -and $NETMASK.split(".").Count -eq 4
            } Until ($isValid -eq $true)
            Do {
                #// Validate DNS server address
                $isValid = $false
                $DNS = Read-Host " Enter your DNS address: "
                $isValid = $DNS -as [ipaddress] -and $DNS.split(".").Count -eq 4
            } Until ($isValid -eq $true)
            Do {
                #// Validate Default Gateway
                $isValid = $false
                $DefaultGw = Read-Host " Enter your Default Gateway: "
                $isValid = $ip -as [ipaddress] -and $DefaultGw.split(".").Count -eq 4
            } Until ($isValid -eq $true)

            #// Set the IP address/Subnet/Gateway of the selected interface to the entered values
            Set-NetIPInterface $selectedIF -AddressFamily IPv4 -Dhcp Disabled

            netsh interface ip set address $selectedIF static $IP $NETMASK $DefaultGw
            #// Set the DNS for the selected interfaces as well
            Set-DnsClientServerAddress -interfaceAlias $selectedIF -serverAddress $DNS
        }
        #// Print the new ipconfiguration
        ipconfig /all
        #// Pause so that the assessor can read the outcome
        PauseMenu
    }
}
#endregion

########################################################################################################
#function to set MacAddress
########################################################################################################
#region MAC Config
function MACMenu($selectedIF)
{
    if(-not $selectedIF) {
        $selectedIF = IFSelectionMenu("MAC Address Configuration")
    }
    #valid macaddress formats are: 00:00:00:11:22:33, 00-00-00-44-55-66, 00.00.00.11.22.33, 000000778899, 000.111.222.333
   [Regex]$validMac = "(?:(?:[a-fA-F0-9]{2}[:\-.]?){5}[a-fA-F0-9]{2}|(?:[a-fA-F0-9]{3}\.){3}[a-fA-F0-9])"
    if($selectedIF -ne "" -and $selectedIF -ne "Back" -and $selectedIF -ne "Skip...") {
        #// Loop input prompt until satisfies MAC address class and length parameters
        Do {
               $macAddress = Read-Host -Prompt "`n Enter the new MAC Address"
            #// Strip all delimiting characters
            $macAddress = $macAddress -Replace "\.","" -Replace "-","" -Replace ":",""
        } Until ($macAddress -match $validMac -and $macAddress.length -eq 12)
        Set-NetAdapter -Name $selectedIF -MacAddress $macAddress | Restart-NetAdapter
        Get-NetAdapter -Name $selectedIF -IncludeHidden
        PauseMenu
    }
}
#endregion

########################################################################################################
#function to set HostName
########################################################################################################
#region HostName Config
function HostNameMenu($quickSetup)
{
    $computerName = Get-WmiObject -class Win32_ComputerSystem
    Do {
        $newHostName = Read-Host -Prompt "`n Enter the new HostName (leave blank to cancel)"
        if ($newHostName -ne "" -and $newHostName -notmatch "^[\w\d_\-]{2,50}$") {
                                Write-Host -ForegroundColor red "`n You entered an invalid HostName"
                                Write-Host -ForegroundColor red "`n Valid HostName can contain alpha numeric characters and _ or -`n"
                }
    } Until ($newHostName -eq "" -or $newHostName -match "^[\w\d_\-]{2,50}$")
    if ($newHostName -eq "") {
                                Write-Host -ForegroundColor yellow "`n Edit HostName cancelled..."
            Sleep -m 1500
    }
    else {
        $computerName.Rename($newHostName) > $null
        $validResponse = @("Y","Yes","N","No")
        $ans = ""

        While ($validResponse -notcontains $ans) {
            Write-Host "`n A reboot is required for the new hostname to take effect"
            $ans = Read-Host -Prompt "`n Would you like to reboot now? (Y | N)"
            if ($ans -eq "Y" -or $ans -eq "Yes") {
                Restart-Computer -Force
            }
            elseif ($ans -eq "N" -or $ans -eq "No") {
                break
            }
        }
        if ($quickSetup) {
            Write-Host -ForegroundColor Green "`n Quick Setup Configuration Complete!"
            Sleep -m 1500
        }
    }
}
#endregion

########################################################################################################
#function to quickly setup values for all menu options to get a system functioning
########################################################################################################
#region Quick Setup Configuration
function QuickSetup()
{
    FirewallMenu($true)
    $selectedIF = IFSelectionMenu "QuickSetup Interface Options" $true
    IPMenu($selectedIF)
    MACMenu($selectedIF)
    HostNameMenu($true)

}
#endregion

########################################################################################################
#Interface Selection Menu
########################################################################################################
function IFSelectionMenu($title, $quickSetup)
{
    #// Get list of Windows Network Adapter objects
    $nics = Get-WmiObject Win32_NetworkAdapter
    #// Create empty arraylist for temporary storage of NIC names
    $netInterfaces = New-Object System.Collections.ArrayList
    #// Create an arraylist to populate the menu items
    $IFOptions = New-Object System.Collections.ArrayList
    #// loop through the nics found
    for($i = 0; $i -le $nics.Count-1; $i++) {
        #// Pull the "name" from each NIC
        $name = $nics[$i].NetConnectionID
        #// If the name isn't null
        if($name) {
            #// Add to the temporary arraylist of nic names
            [void]$netInterfaces.Add($name)
            #// Add 1 to the integer for quick and dirty hot-key shortcut
            $num = $i+1
            #// Just call "Break()" since only menu selection item matters
            $cmd = "Break"
            #// This adds the individual menu items for each NIC to the arraylist
            [void]$IFOptions.Add(@($name,$cmd,"$num"))
        }
    }
    #// Add the extra "back/exit" option to the end
    Write-Host $quickSetup
    if ($quickSetup) {
        [void]$IFOptions.Add(@("Skip...","Break","X"))
    }
    else {
        [void]$IFOptions.Add(@("Back","Break","X"))
    }
    #// Create the new sub-menu to display the NICs and save the selection to $selectedIF
    return MakeMenu @("$title","Please select an Interface") "70" "+" $IFOptions
}

########################################################################################################
#Script Initiation
########################################################################################################
clear-host
Set-ExecutionPolicy bypass
#// Create and initiate the main menu options
<#MakeMenu "Virtual Machine Setup Script" "70" "+" @( ("Firewall Options","FirewallMenu", "1"),
                                                    ("IP Configuration","IPMenu", "2"),
                                                    ("MAC Configuration","MACMenu","3"),
                                                    ("HostName Configuration","HostNameMenu","4"),
                                                    ("Quick Setup Configuration","QuickSetup","Q"),
                                                    ("Reset VM Configuration","ResetSetup","R"),
                                                    ("Exit Script","Break","X") )#>
MakeMenu "Virtual Machine Setup Script $versionNum" "70" "+" @( ("Firewall Options","FirewallMenu", "1"),
                                                    ("IP Configuration","IPMenu", "2"),
                                                    ("MAC Configuration","MACMenu","3"),
                                                    ("HostName Configuration","HostNameMenu","4"),
                                                    ("Quick Setup Configuration","QuickSetup","Q"),
                                                    ("Exit Script","Break","X") )
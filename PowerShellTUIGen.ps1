$menuSelectionIndex = 0
$moduleTable = @{}


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


function PauseMenu()
{
    $Global:HostConnectCheck = 0
    Write-Host "`n Press any key to continue..."
    #// Start key input verification, this DOES NOT WORK inside of the PowerShell ISE
    $keyInput = $host.UI.RawUI.ReadKey("NoEcho, IncludeKeyDown")
}


#// This is the module loader, that parses, loads the "module" scripts into memory, unloads the ones that aren't valid
function BlueLite_LoadModules($folder = "WMI")
{
    $menuItems = New-Object System.Collections.ArrayList
    $moduleTable.Clear
    $countGood = 0
    $automateParam = $false #// We can come back to this later, not worth the time for now...
    Get-ChildItem -Recurse "$pwd\$folder" | ForEach-Object {
        if($_.Extension -eq ".psm1") {
            Import-Module $_.FullName -WarningAction SilentlyContinue -Force
            if(Test-Path function:/MenuInterface) {
                if($automateParam) {
                    $paramList = (Get-Command -Name $_.BaseName).Parameters
                    forEach($var in $paramList.Keys) {
                        if($paramList[$var].Attributes.TransformNullOptionalParameters) {
                            "----------------------------------------------"
                            $paramList[$var].Name
                            $paramList[$var].Attributes
                            $paramList[$var].Values
                            #$paramList[$var].ParameterType
                            "----------------------------------------------"
                            Write-Host -f Green -b DarkGreen $paramList[$var].Name "; " $paramList[$var].ParameterType "; " $paramList[$var].Values
                            PauseMenu
                        }
                    }
                }
                else {
                    $moduleData = MenuInterface
                    Set-Item function:/MenuInterface {} -Force
                    if($moduleData.count -ne 0) {
                        if($moduleData["type"] -eq $folder) {
                            $countGood++
                            $moduleName = $_.BaseName
                            $moduleTable[$moduleName] = $moduleData
                            $moduleCall = "BlueLite_ParamMenu(""$moduleName"")"
                            [void]$menuItems.Add(@($moduleData["title"],"$moduleCall",[string]$countGood))
                        }
                        else {
                            Remove-Module $_.BaseName -WarningAction SilentlyContinue
                        }
                    }
                }
            }
        }
    }
    if($countGood -gt 0) {
        [void]$menuItems.Add(@("Back","Break","X"))
        MakeMenu @($folder,"Select a Script") "70" "+" $menuItems
    }
    else {
        Write-Host -f Yellow -b DarkYellow "`n No Modules could be found!"
        PauseMenu
    }
}


#// This builds the list of parameters available to that script/function, and adds the execute function as well to execute the function based on parameters set or defaults
function BlueLite_ParamMenu($moduleName)
{
    $moduleData = $moduleTable[$moduleName]
    $menuItems = New-Object System.Collections.ArrayList
    #$execCall = $moduleData["function"] + " " #// This is the execution call that will compile the variables if they exist, if not the defaults
    #$execCall = $moduleName #// This is the execution call that will compile the variables if they exist, if not the defaults
    for($i = 0; $i -lt $moduleData["params"].count; $i++) {
        $param = $moduleData["params"][$i]
        $paramCall = "BlueLite_ParamMenuEdit """ + $moduleName + """ """ + ($i) + """"
        [void]$menuItems.Add(@($param["varDesc"],$paramCall,[string]($i+1)))
    }
    $execCall = "BlueLite_ExecFunc " + $moduleName
    [void]$menuItems.Add(@("Execute Script",$execCall,"E"))
    [void]$menuItems.Add(@("Back","Break","X"))
    MakeMenu @($moduleData["title"],"Script Parameters") "70" "+" $menuItems
}


#// This rebuilds the execution line of the function based on the currently set values, or defaults if none are new
function BlueLite_ExecFunc($moduleName)
{
    $moduleData = $moduleTable[$moduleName]
    $execCall = $moduleName
    for($i = 0; $i -lt $moduleData["params"].count; $i++) {
        $param = $moduleData["params"][$i]
        if($param["newVal"]) { #// Simply check if this extra hashtable entry exists, if there was manual entry, if it does use it!
            $varVal = $param["newVal"]                                    
        }
        else {
            $varVal = $param["defVal"]
        }
        
        if($param["varType"] -eq "Switch") {
            #// It's the thing that causes us so many issues
            if($varVal -eq "True") { #// Basically this just ensures that the only time a "Switch" is added to the execution is if it's set to "True"
                $execCall += " -" + $param["varName"]
            }
        }
        elseif($param["varType"].Contains("[]")) {
            #// It's an array, probably have to add some "()" or "," or something
            $execCall +=  " -" + $param["varName"] + " " + $varVal
        }
        else {
            $execCall +=  " -" + $param["varName"] + " " + $varVal
        }
    }
    $execCall += ";PauseMenu" #// TODO: Instead of PauseMenu, we should have a special pause here with the option to save/export the output!!
    Write-Host $execCall
    Invoke-Expression $execCall
}


function BlueLite_ParamMenuEdit($moduleName, $paramIndex)
{
    $param = $moduleTable[$moduleName]["params"][$paramIndex] #this is an array of the parameters in this given module's array of parameters...
    
    #Loop until back is selected (can't preserve the menuSelection this way, but it works)
    Do {
        if($param["newVal"] -eq "" -or $param["newVal"] -eq $null) {
            $param["newVal"] = $param["defVal"]
        }


        $selection = MakeMenu "Editing $moduleName" "70" "+" @(@("$($param[""varName""]) = $($param[""newVal""])", "break", 1),
                                                               @("Use default value ($($param[""defVal""]))",'$param["newVal"] = $param["defVal"];break',"d"),
                                                               @("Back", "break", "x"))


        #if there is an equals symbol then prompt for the new value
        if($selection.Contains(" = "))
        {
            if($param["varType"] -eq "String[]") {
                $input = Read-Host "Keep existing values (y/n)?"
                if($input -eq "y" -or $input -eq "yes") {
                    $tempArray = $param["newVal"]
                } else {
                    $tempArray = @()
                }


                Do {
                    $input = Read-Host "Input value to add (enter q to quit)"
                    if($input -ne "q") {
                        $tempArray += $input
                    }
                } Until($input -eq "q")


                $param["newVal"] = $tempArray
            } elseif($param["varType"] -eq "Switch") {
                $input = Read-Host "`nEnable the switch (y/n)?"
                if($input -eq "y" -or $input -eq "yes") {
                   $param["newVal"] = "y"
                } elseif($input -eq "n" -or $input -eq "no") {
                   $param["newVal"] = "n"
                }
            } elseif($param["varType"] -eq "Int") {
                $param["newVal"] = Read-Host "`nEnter new value"
            } elseif($param["varType"] -eq "String") {
                $param["newVal"] = Read-Host "`nEnter new value"
            }
        }
    } Until($selection -eq "Back")
}


#// TODO :: This will be the custom editor for TargetList variables since it's such a common thing
function BlueLite_TargetList()
{
}


#// TODO :: This will be where we configure/establish the config.ini, but it won't break like the old one, and maybe store some persistence for script defaults too?
function BlueLite_Config()
{
}


########################################################################################################
#Script Initiation
########################################################################################################
clear-host


MakeMenu @(":::::::BlueLite:::::::","BlueLigHT lightweight interface") "70" "+" @(  ("WMI (Laser)",'BlueLite_LoadModules("WMI")', "1"),
                                                                                    ("WinRM (Torch)",'BlueLite_LoadModules("WinRM")', "2"),
                                                                                    ("Manage Target List","BlueLite_TargetList", "3"),
                                                                                    ("Configuration","BlueLite_Config", "4"),
                                                                                    ("Exit Script","Break","X") )
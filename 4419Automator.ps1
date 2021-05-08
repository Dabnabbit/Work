#region MainForm
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
#region Classes
Class Template{
    [String]$templateName
    [String]$TemplateFile
    [String]$RosterFile
    [String]$GradeBookFile
    [String]$SaveDir
    [String]$Instructor
    [String]$Role
    [String]$StartDate
    [String]$EndDate
    [String]$StrCourseDescription
    [String]$StrDemonstrated
    [String]$StrAcademicExcellence
    [String]$StrDistinguishedGraduate
    [String]$StrOutstandingContributor

}
Class Student{
  [String]$StudentID
  [String]$LastName
  [String]$FirstName
  [String]$Rank
  [String]$Unit
  [bool]$Fails
  [System.Collections.Specialized.OrderedDictionary]$Grades
}
#endregion
#region Variables
[String]$global:AppTitle = "Automated 4419 Creator"

#//-Need to work on making these customizable through the GUI for different courses...
[String]$global:strDEFAULTCourseDescription = "--%rank %lastname attended a 26 day CVA/Hunt Weapon System Training course for Cyberspace Operator pre-certification from %datestart - %datestop. This training included CVAH Weapons System Configuration/Operation and Cyber Protection Team Methodologies. They completed all knowledge and task requirements for course completion. The training involved a combination of hands-on Performance Evaluations and cognitive Academic Evaluations; scores on reverse.`n`n(Continued on reverse)"
[String]$global:strDEFAULTDemonstrated = "--%rank %lastname demonstrated a solid understanding of all topics covered throughout the course. Recommend expediting student training to a level commensurate with their skill and experience"
[String]$global:strDEFAULTAcademicExcellence = "--Awarded Academic Excellence based on having a GPA above 95%"
[String]$global:strDEFAULTDistinguishedGraduate = "--Awarded Distinguished Graduate based on having above a 95% GPA, being top 10% of class, and displaying character deserving of award."
[String]$global:strDEFAULTOutstandingContributor = "--Awarded Outstanding Contributor based on student selection of classmate who had the most positive impact on other students learning."
[String]$filePath
[String]$templateName
[String]$gradesFileName
[String]$rosterFileName
[String]$Instructor
[String]$strCrewPosition
[String]$strStartdate
[String]$strGradDate
[template]$global:currentTemplate = New-Object -TypeName Template
[template[]]$global:templates = @()
[student[]]$global:Students = @()
#endregion

$global:AppColorBG = [System.Drawing.Color]::FromArgb(250,250,250)
$global:AppColorTrim = [System.Drawing.Color]::FromArgb(255,255,255)

#//-Font/Format normalization globals
$global:font0 = New-Object System.Drawing.Font('Calibri',10) #//- Inputs
$global:font1 = New-Object System.Drawing.Font('Calibri',14,'Bold','Pixel') #//- Title, Labels, "...", "Process", "Close", Dropdown selection
$global:font2 = New-Object System.Drawing.Font('Calibri',18,'Bold','Pixel') #//- "+", "-","..."
$global:font3 = New-Object System.Drawing.Font('Calibri',14,'Bold','Pixel') #//- "Process" / "Close"
$global:font4 = New-Object System.Drawing.Font('Calibri',14,'Bold','Pixel') #//- Console Output
#region Helper Functions
function Get-Property {
    param([__ComObject] $object, [String] $propertyName)
    $obj = [System.Type]::GetType($object)
    $obj.InvokeMember($propertyName, "GetProperty",
    $NULL, $object, $NULL)
}
#-Set-Property--------------------------------------------------------
function Set-Property {
    param([__ComObject] $object, [String] $propertyName,
    $propertyValue)
    $obj = [System.Type]::GetType($object)
    [Void] $obj.InvokeMember($propertyName, "SetProperty",
    $NULL, $object, $propertyValue)
}
#-Invoke-Method-------------------------------------------------------
function Invoke-Method {
    param([__ComObject] $object, [String] $methodName,
    $methodParameters)
    $obj = [System.Type]::GetType($object)
    $output = $obj.InvokeMember($methodName, "InvokeMethod",
    $NULL, $object, $methodParameters)
    if ( $output ) { $output }
}
function UpdateText([String] $Caption, [int] $curNum, [int]$maxNum){
  [System.Drawing.Graphics]$gr = $progress.CreateGraphics()
  $global:progress.Value = [math]::Ceiling(($curNum /$maxNum) * 100)
  $x = ($global:progress.Width / 2) - ([int]$gr.MeasureString($Caption, $global:Font).Width / 2)
  $y = ($global:progress.Height / 2) - ([int]$gr.MeasureString($Caption, $global:Font).Height / 2)
  $PointF = [System.Drawing.PointF]::new($x, $y)
  # $Caption = "$($global:progress.Value) % Complete"
  $gr.DrawString($Caption, $global:Font, $global:brush1, $PointF)
}
function SaveSettings() {
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
  $global:templates | Export-Csv $ScriptDir\templates.csv
}
function LoadSettings {
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
  If (Test-Path $ScriptDir\templates.csv){
    Write-Info -txt "Templates Loaded"
    $global:templates = Import-Csv $ScriptDir\templates.csv
    $cboTemplate.Items.Clear()
    $cboTemplate.Items.AddRange($global:templates.TemplateName)
    $cboTemplate.SelectedIndex = 0
  }
}
function DeleteTemplate {
  param (
    $templateName
  )
  $global:templates = $global:templates | Where-Object {$_.templateName -ne $templateName}
  if ($cboTemplate.Items.IndexOf($templateName) -eq 0){
    $cboTemplate.Items.RemoveAt($cboTemplate.SelectedIndex)
    $cboTemplate.SelectedIndex = 0
  }else{
    $tmp = $cboTemplate.SelectedIndex
    $cboTemplate.Items.RemoveAt($tmp)
    $cboTemplate.SelectedIndex = ($tmp - 1)
  }
  $global:templates | Export-Csv .\templates.csv
}
function UpdatePDF {
  [CmdletBinding()]
  Param
  (
      $filePath = $global:txtSaveDir.text,
      $templateName = $global:txtTemp.text,
      $gradesFileName = $global:txtGradeBook.text,
      $rosterFileName = $global:txtRoster.text,
      $Instructor = $global:txtInstructor.text,
      $strCrewPosition = $global:txtRole.text,
      $strStartdate = $global:StartDate.Text,
      $strGradDate = $global:EndDate.Text
  )
  Begin
  {
    $global:progress.visible = $true
  }
  Process
  {
    #$DebugPreference = "Continue"
    $ReturnCode = ProcessRoster
    if ( $ReturnCode -ne "Error"){
      ProcessGradeBook
      WritePDF
    }
  }
  End
  {
    $global:progress.Visible = $false
    Write-Info -OutputType "Info" -txt "[ * ] Process Complete"
  }
}
function WritePDF {
  Begin{
    try {
      $AcroApp = New-Object -ComObject AcroExch.App
      $theForm = New-Object -ComObject AcroExch.PDDoc
      $jso = [object]
      #$AcroApp.Hide()
    }
    catch {
      Write-Info -OutputType "Error" -txt "[ - ] Adobe is not Installed"

    }

  }
  Process{
    if ($theForm){
      if (!$theForm.Open($global:txtTemp.text)){
        [System.Windows.MessageBox]::Show("Can't Open Template")
        return
      }
      $jso = $theForm.GetJSObject()
      #$AVDoc = $theForm.OpenAVDoc("")
      if ($jso){
        $xfa = Get-Property $jso "xfa" @(0)
        [int]$x = 0
        if ($xfa){
          Write-Info -OutputType "Info" -txt "[ * ] Preparing Forms..."
          foreach ($tmpStudent in $global:Students) {
            UpdateText -curNum $x -maxNum $global:Students.count  -Caption "Writing PDF Files"
            #Write-Info -OutputType "Info" -txt "[ * ] Preparing Forms for $($tmpstudent.Rank) $($tmpStudent.LastName)"
            #Front Page
            $node = Invoke-Method $xfa "ResolveNode" @("form.form1.page1.txtUnit")
            Set-Property $node "formattedValue" @($tmpstudent.Unit)
            $node = Invoke-Method $xfa "ResolveNode" @("form.form1.page1.txtStudent")
            Set-Property $node "formattedValue" @("$($tmpstudent.Rank) $($tmpstudent.FirstName) $($tmpstudent.LastName), $($global:txtRole.text)")
            $node = Invoke-Method $xfa "ResolveNode" @("form.form1.page1.dateDate")
            Set-Property $node "formattedValue" @($((get-date $strGradDate -Format "yyyy-MM-DD")).ToString())
            $node = Invoke-Method $xfa "ResolveNode" @("form.form1.page1.txtInstructor")
            Set-Property $node "formattedValue" @($global:txtInstructor.Text)
            $node = Invoke-Method $xfa "ResolveNode" @("form.form1.page1.txtInstComments")
            #////Set-Property $node "formattedValue" @((((($strCourseDescription -replace "%rank", $tmpstudent.Rank) -replace "%lastname", $tmpstudent.LastName) -replace "%days", "26" ) -replace "%datestart", $strStartdate ) -replace "%datestop", $strGradDate)
            #//-Trying to utilize the new text field
            Set-Property $node "formattedValue" @((((($global:txtStrCourseDescription.Text -replace "%rank", $tmpstudent.Rank) -replace "%lastname", $tmpstudent.LastName) -replace "%days", "26" ) -replace "%datestart", $strStartdate ) -replace "%datestop", $strGradDate)
            #Back Page
            $dictClass = @{}
            $tmpGPA = 0
            foreach ($modBlock in $tmpStudent.Grades.Keys) {
              $tmpMod,$tmpBlock = $modBlock -split "::"
              $tmpGrade = $tmpStudent.Grades[$modBlock]
              $tmpSubLine = "-- $tmpBlock`: $tmpGrade%`n"
              if ($tmpBlock){
                if ($dictClass.keys.Count -eq 0){
                  $dictClass[$tmpMod]=$tmpSubLine
                }else{
                  $dictClass[$tmpMod] = "$($dictClass[$tmpMod])$tmpSubLine"
                }
              }elseif ($TmpMod -like "*GPA*") {
                $dictClass["GPA"] = "$tmpGrade%"
                $tmpGPA = [double]$tmpGrade
              }
            } #End foreach grade keys
            $strGradeBlock = ""
            #//-Now let's run through that temporary dictionary and generate the bulk of the gradeblock for each class
            foreach ($modType in $dictClass.Keys) {
              if ($modType -eq "GPA"){
                $strGradeBlock = "Overall GPA: $($dictClass[$modType])`n$strGradeBlock"
              }else{
                $strGradeBlock += "`n$modType`n$($dictClass[$modType])"
              }
            }
            #//-Academic Excellence, make this an input parameter or something that can be referenced easier?
            If ($tmpGPA -gt 95 -and $tmpstudent.Fails -ne $true) {
              $strGradeBlock = "$strGradeBlock`n$($global:txtStrAcademicExcellence.Text)`n"
            }
            #//-Distinguished Graduate, make this an input parameter or something that can be referenced easier?
            #If tmpGPA > 95 And tmpstudent.Fails <> True Then
            #    strGradeBlock = strGradeBlock & vbNewLine & strDistinguishedGraduate & vbNewLine
            #End If
            #//-Outstanding Contributor, make this an input parameter or something that can be referenced easier?
            # Not implemented yet! **use strOutstandingContributor for verbage**
            #If tmpGPA > 95 And tmpstudent.Fails <> True Then
            #    strGradeBlock = strGradeBlock & vbNewLine & strOutstandingContributor & vbNewLine
            #End If
            #$strGradeBlock = "$strGradeBlock`n$(-Replace(-Replace($strDemonstrated, "%rank", $tmpstudent.Rank), "%lastname", $tmpstudent.LastName))`n"


            #$strGradeBlock = ("$strGradeBlock`n$(($strDemonstrated -replace '%rank', $tmpstudent.rank) -replace '%lastname',$tmpstudent.LastName)`n")
            $strGradeBlock = ("$strGradeBlock`n$(($global:txtStrDemonstrated.Text -replace '%rank', $tmpstudent.rank) -replace '%lastname',$tmpstudent.LastName)`n")


            Write-Debug $strGradeBlock
            $node = Invoke-Method $xfa "ResolveNode" @("form.form1.page2.txtDescription")
            Set-Property $node "formattedValue" @($strGradeBlock)
            #$jso.xfa.ResolveNode("form.form1.page2.txtDescription").formattedValue = $strGradeBlock
            #//-Save and close the PDF file, generating the new filename based on filling in the data from the template's filename
            If ($tmpstudent.FirstName) {
                $fileName = Split-Path $global:txtTemp.Text -Leaf
                $theForm.Save(1, "$($global:txtSaveDir.text)\$(((($fileName -Replace "LastName", $tmpstudent.LastName) -replace "FirstName", $tmpstudent.FirstName) -replace "Rank", $tmpstudent.Rank) -replace "Unit", $tmpstudent.Unit)")
                $node = Invoke-Method $xfa "ResolveNode" @("form.form1.page1.txtStudent")
                Write-Debug "Finished: $(Get-Property $node "FormattedValue" @(0))"
            }
            $x++
          } # End foreach student
        } # End Checking XFA
      } # End Checking jso
    }
  } #end process block
  End{
    if ($AcroApp){
      $AcroApp.CloseAllDocs()
      $AcroApp.exit()
    }
  }
}
function ProcessRoster {
  Begin{
    try {
      $excel = New-Object -Com Excel.Application
    }
    catch {
      Write-Info -OutputType "Error" -txt "[ - ] Excel is not Installed"
    }

  }
  Process{
    if ($excel){
      #//-Parse through Roster file generating student data for all attending students
      $excel.displayAlerts = $false
      $wbookRoster = $Excel.Workbooks.Open($global:txtRoster.text, $false)
      if ($wbookRoster){
        $sheetRoster = $wbookRoster.Worksheets(1)
        $lastrowroster = $sheetRoster.UsedRange.Rows.Count
        $lastColumnRoster = $sheetRoster.UsedRange.Columns.Count
        $HeaderColumnRoster = new-object hashtable
        $HeaderColumnRoster.Add("Last Name", 0)
        $HeaderColumnRoster.Add("First Name", 0)
        $HeaderColumnRoster.Add("Rank/      Grade", 0)
        $HeaderColumnRoster.Add("Unit", 0)
        $intHeaderMax = $HeaderColumnRoster.Count
        $intHeaderCnt = 0
        $intHeaderEnd = 0
                                $OffSet = 0
        Write-Info -OutputType "Info" -txt "[ * ] Analyzing Roster"
        for ($i = 1; $i -le $lastrowroster; $i++){
          $rowRoster = $sheetRoster.Rows($i)
          if (($rowRoster.value2 -join "") -ne ""){Write-Debug ($rowRoster.value2 -join "")}
          for ($a = 1; $a -lt $lastColumnRoster; $a++){
              if($sheetRoster.cells($i,$a).value2){
                Write-Debug "Current Offset: $OffSet"
                if ($OffSet -eq 0){
                    $offset = $a
                    break
                } # End If Offset = 0
              } # End Checking for Value
          } # End for Loop Offset
                                  if ($sheetRoster.cells($i,($offset)).value2){
            if (($sheetRoster.Cells($i, ($offset)).Value2).trim()){
              for ($j = $OffSet; $j -le $lastColumnRoster; $j++){
                $cellRoster = ($sheetRoster.Cells($i, $j).Value2)
                if ($cellRoster) {$cellRoster = $cellRoster.trim()}
                if ($cellRoster -like "*UNCLASSIFIED//*"){break}
                If ($cellRoster){
                  #//-Chop off anything after a comma in the header, this helps resolve the middle initial issue
                  $tmpHeadComma = ($cellRoster -split ",")[0]
                  if ($tmpHeadComma -in $HeaderColumnRoster.Keys){
                    $HeaderColumnRoster[$tmpHeadComma] = $j
                    $intHeaderCnt++
                    if ($intHeaderCnt -eq $intHeaderMax){break}
                  } # End If Keys
                } # End CellRoster
            } # End For LastColumnRoster

            If ($intHeaderCnt -eq $intHeaderMax){
            $intHeaderEnd = $i
              break
            } # End If
            } # End SheetRoster Cell
          } # End Checking for Null
        } # End For Roster Row
      if($intHeaderEnd -lt 1){
        #Write-Info -OutputType "Error" -txt "[ - ] Error Finding Headers in Roster"
        Write-Info -OutputType "Error" -txt "[ - ] Could not find Header(s): $(($headerColumnRoster.Keys | Where-Object {$headerColumnRoster[$_] -eq 0}) -join " ")"
        return "Error"
      }
      Write-Info -OutputType "Info" -txt "[ * ] Located Student Data"
        for ($i = $intHeaderEnd +1; $i -le $lastrowroster; $i++){
          #Write-Debug $i
          UpdateText -curNum $i -maxNum $lastrowroster -Caption "Processing Roster"
          #$rowRoster = $sheetRoster.Rows($i)
          if ($sheetRoster.Cells($i, $OffSet).Value2){
            If (($sheetRoster.Cells($i, $OffSet).Value2).trim() -And -Not $sheetRoster.Cells($i, $OffSet).MergeCells){
              $tempStudent = ""
                #$dictStudent = New-Object hashtable
                For ($j = $OffSet; $j -le $lastColumnRoster; $j++){
                  if ($sheetRoster.Cells($i, $j).Value2){
                    $cellRoster = (($sheetRoster.Cells($i, $j).Value2).tostring()).trim()
                   If ($cellRoster){
                        #Debug.Print "Is student data:", cellRoster
                        $tempStudent = "$tempStudent$j`:$($cellRoster -Replace ',', ''),"
                    }
                  } # End Checking for Null
                } # End For J Loop
                If ($tempStudent){
                    #//-Chop off the trailing comma seperator
                    $tempStudent = $tempStudent.Substring(0,($tempStudent.Length - 1))
                    write-Debug "STUDENT: $tempStudent"
                    #//-Break data back out into an array with column number and entry
                    $dataStudent = $tempStudent -split ","
                    $tStudent = New-Object -TypeName Student
                    foreach ($data in $dataStudent) {
                      #//-Split by : and then find Header names from column numbers for each entry
                      $tmpSplit = $data -split ":"
                      $tmpHeadr = ""
                      foreach ($tmpKey in $HeaderColumnRoster.Keys) {
                        #Debug.Print tmpKey, HeaderColumnRoster(tmpKey), tmpSplit(0), (CStr(HeaderColumnRoster(tmpKey)) = CStr(tmpSplit(0)))
                        If ([string]$HeaderColumnRoster[$tmpKey] -eq [string]$tmpSplit[0]) {
                          $tmpHeadr = $tmpKey
                      }
                      } # End headercolumnRoster.keys
                      #Debug.Print tmpHeadr, tmpSplit(1) #add
                      if ($TmpHeadr -ne ""){
                        switch ($tmpHeadr)
                          {
                              "Last Name" {$tStudent.LastName = $tmpSplit[1];break}
                              "First Name" {
                                #//-Check if there is a MI in the Name, then remove it
                                if ($tmpSplit[1].IndexOf(".") -gt 0){
                                  $tmpSplit[1] = ($tmpSplit[1].Substring(0,($tmpSplit[1].Length -2))).trim() #Trim(Mid(tmpSplit(1), 1, (Len(tmpSplit(1)) - 2)))
                                } # End If "."
                                $tStudent.FirstName = $tmpSplit[1]
                                break
                              } # End First Name
                              "Rank/      Grade" {$tStudent.Rank = $tmpSplit[1];break}
                              "Unit" {$tStudent.Unit = $tmpSplit[1];break}
                          } # End Switch
                      } # End If TmpHeader <> ""
                      #Debug.Print tmpHeadr, tmpSplit(1)
                    }
                    #//-Create entry for that student
                    $global:Students += $tStudent
                  }
            } # End If
          } # End checking if value
          #//-Space after finding several students indicates students no longer enrolled on roster, revisit this later if info is needed from disenrolled
          ElseIf ($Students.Count -gt 0) {
            #//-Terminate loop through rows once we find the first blank line AFTER finding students
            UpdateText -curNum $lastrowroster -maxNum $lastrowroster -Caption "Processing Roster"
            Write-Info -OutputType "Info" -txt "[ * ] Found $($global:Students.Count) Students"
            break
          } # End Elseif Student.count
        } # End For Loop
        return "OK"
      } # End Checking for wbookRoster
    } # End Checking for Excel
  } # End Process Block
  End{
    if ($wbookRoster){$wbookRoster.close($false)}
    if ($excel){$excel.quit()}

  }
}
function ProcessGradeBook {
  Begin{
    try {
      $excel = New-Object -Com Excel.Application
    }
    catch {
      Write-Info -OutputType "Error" -txt "[ - ] Excel is not Installed"
    }

  }
  Process{
    if ($excel){
      $wbookGrades = $Excel.Workbooks.Open($global:txtGradeBook.text)
      if ($wbookGrades){
        $sheetGrades = $wbookGrades.Worksheets(1)
        $lastRowGrades = $sheetGrades.UsedRange.Rows.Count
        $lastColumnGrades = $sheetGrades.UsedRange.Columns.Count
        #//-Next, Parse through the Gradebook excel file looking for students from Roster
        Write-Info -OutputType "Info" -txt "[ * ] Processing GradeBook"
        $rowHeader = 0
        for ($i = 1; $i -lt $lastRowGrades; $i++){
          UpdateText -curNum $i -maxNum $lastRowGrades -Caption "Processing Gradebook"
            if ($sheetGrades.cells($i,1).value2){
              If (($sheetGrades.Cells($i, 1).Value2.toString()).trim()) {
                  #//-Is header row?
                  Write-Debug ($sheetGrades.Cells($i, 1).Value2.toString()).trim()
                  If (($sheetGrades.Cells($i, 1).Value2.toString()).trim() -eq "#") {
                      Write-Debug "Is header row $i"
                      $rowHeader = $i
                      #//-I don't like this hack, but it should skip the next row (because of how the headers are used) to avoid detecting end of gradesheet
                      $i++
                  } #end if
                  ElseIf ($rowHeader -gt 0){
                      $LastName = ""
                      $FirstName = ""
                      $tStudent = New-Object -TypeName Student
                      #//-Start at 2 to avoid the #
                      for ($j = 2; $j -le $lastColumnGrades; $j++) {
                        #//-This should help handle any merged cell headers
                        if ($sheetGrades.Cells($rowHeader, $j).MergeArea.Cells(1, 1).Value2){
                          $cellMod = ($sheetGrades.Cells($rowHeader, $j).MergeArea.Cells(1, 1).Value2).trim()
                          #//-Exit loop if we find a blank column, extra data we don't need after this
                          If (!$cellMod){
                            break
                          }
                          #//-Bit Hacky but again, because of how the header rows are setup this grabs the secondary header row
                          if ($sheetGrades.Cells($rowHeader + 1, $j).MergeArea.Cells(1, 1).Value2){
                            $cellBlock = ($sheetGrades.Cells($rowHeader + 1, $j).MergeArea.Cells(1, 1).Value2.toString()).trim()
                            if(($sheetGrades.Cells($i, $j).Value2)){
                              $cellGrades = ($sheetGrades.Cells($i, $j).Value2.toString()).trim()
                              #//-Remove redundancy in non-class headers
                              write-Debug (($j, $cellMod, $cellBlock, $cellGrades) -join " ")
                              If ($cellMod -eq $cellBlock){
                                  $cellBlock = ""
                              }
                              switch ($cellMod) {
                                "LAST" { $LastName = $cellGrades ;break}
                                "FIRST" {
                                  $FirstName = $cellGrades
                                  If ($cellGrades.IndexOf(".") -gt 0 ){
                                      $FirstName = ($cellGrades.Substring(0,($cellGrades.Length -2))).trim()
                                      break
                                 } }
                                Default {
                                  If ($LastName -And $FirstName) {
                                          If (!$tStudent.LastName){
                                            #//-Loop through Roster of students
                                            foreach ($tmpstudent in $global:students) {
                                              #Write-Debug (($LastName, $FirstName, $tmpstudent.LastName, $tmpstudent.FirstName) -join ",")
                                              If ($tmpstudent.LastName -eq $LastName -and $tmpstudent.FirstName -eq $FirstName) {
                                                  write-Debug "Found them! $($tmpstudent.LastName), $($tmpstudent.FirstName)"
                                                  $tStudent = $tmpstudent
                                                  $tStudent.Grades = New-Object System.Collections.Specialized.OrderedDictionary #@{}
                                                  break
                                              } # End If
                                            } # End Foreach Student
                                            If (!$tStudent.LastName) {
                                              #//-Trigger this block of code if a student in the Roster is missing from the Gradebook!
                                              Write-Debug "Couldn't find Student: $LastName, $FirstName"
                                              Write-Info -OutputType "Error" -txt "[ - ] No Roster Entry for $lastName, $FirstName"
                                              $global:Students = $global:students | Where-Object {$_.FirstName -ne $FirstName -and $_.LastName -ne $LastName}
                                              $LastName = $null
                                              break
                                            } # End Missing
                                          } # End If No Student.LastName
                                          #else{
                                            #//-Ignore any Mod Total scores, only need the individuals and the overall GPA
                                            If ($cellBlock -ne "Total") {
                                              #Debug.Print i, j, cellBlock
                                              #//-We should now have the matched tStudent roster object, and this is each column of grade data
                                              $modBlock = $cellMod
                                              #//-To keep it a simple dictionary, gonna combine mod/block into one with a "::" separator to be split later
                                              #If cellBlock <> "" Then
                                              $modBlock = $modBlock + "::" + $cellBlock
                                              #End If
                                              #//-Add grade/row entry into Student.Grades dictionary, and format Grade to two decimal places
                                              $tStudent.Grades[$modBlock] = '{0:N2}' -f [double]$cellGrades
                                              #//-Check for comments on grade blocks, if there is a comment then it was a fail, and disqualifies from Academic Excellence
                                              If (($sheetGrades.Cells($i, $j)).Comment){
                                                  #Debug.Print "Uh oh, a fail :(", tStudent.LastName, tStudent.FirstName, modBlock, sheetGrades.Cells(i, j).Comment.Text
                                                  $tStudent.Fails = $true
                                              } # End If
                                              Write-Debug ($tStudent.LastName, $tStudent.FirstName, $modBlock, $tStudent.Grades[$modBlock] -join " ")
                                            } # End If
                                        #}
                                  } # End If LastName and FirstName
                                } # End Default
                              } # End Switch
                            } # End Checking for Null
                          } # End Checking for Null
                        } # End Checking for Null
                      } # End J For Loop
                      # Added Grades to student Records
                      for ($ii = 0; $ii -lt $global:students.Count; $ii++) {
                        if (($global:students[$ii].LastName -eq $tStudent.LastName) -and $global:students[$ii].FirstName -eq $tStudent.FirstName){
                          $global:students[$ii].Grades = $tStudent.Grades
                          break
                        }
                      } # End For Loop
                  } # End ElseIf
              } # End If
            } # End Checking for Null
            #//-As long as we have headers, blank row will indicate end of grades
            #//    Note: the current Gradebook has a value of "1" in the first column of the end...
            #//    this causes it to pick up a few extra bits of data
            ElseIf ($rowHeader -gt 0) {
              UpdateText -curNum $lastRowGrades -maxNum $lastRowGrades -Caption "Processing Gradebook"
                Write-Debug "End of GradeBook Detected at line: $i"
                break
            }
        } # End For loop
      } # End Checking for wbookGrades
    } # End Checking for Excel
  } # End Process Block
  End{
    $wbookGrades.close($false)
    $excel.quit()
  }
}
function Write-Info  {
  param (
    [String]$txt,
    [String]$OutputType
  )

  switch ($OutputType) {
    "Error" { $global:rtbOutput.SelectionColor='Red' }
    "Info"      { $global:rtbOutput.SelectionColor='Green' }
    Default { $global:rtbOutput.SelectionColor='White'}
  }

  $global:rtbOutput.AppendText("[$(Get-date -Format "hh:mm:ss")] $txt`n")
  $global:rtbOutput.ScrollToCaret()
}
#endregion
#region Form Creation
#region Main Form
$frmMain = New-Object system.Windows.Forms.Form -Property @{
    ClientSize = New-Object System.Drawing.Point(1077,589)
    font = $global:font1
    text = $global:AppTitle
    TopMost = $false
    MaximizeBox = $false
    MinimumSize = "1077,589"
    BackColor = $global:AppColorBG
}
#endregion

$tblTrim = New-Object System.Windows.Forms.TableLayoutPanel -Property @{
    ColumnCount = 3
    RowCount = 3
    Dock = "Fill"
    Anchor = "None"
    BackColor = $global:AppColorTrim
}
$tblTrim.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 5))) | Out-Null
$tblTrim.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 90))) | Out-Null
$tblTrim.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 5))) | Out-Null
$tblTrim.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,5))) | Out-Null
$tblTrim.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,90))) | Out-Null
$tblTrim.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,5))) | Out-Null

#region Main TableLayout
$tblMain = New-Object System.Windows.Forms.TableLayoutPanel -Property @{
    ColumnCount = 6
    RowCount = 13
    Dock = "Fill"
    Anchor = "None"
    BackColor = $global:AppColorBG
}
$rowHeight = 8

$tblMain.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 15))) | Out-Null
$tblMain.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, $rowHeight))) | Out-Null
$tblMain.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, $rowHeight))) | Out-Null
$tblMain.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, $rowHeight))) | Out-Null
$tblMain.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, $rowHeight))) | Out-Null
$tblMain.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, $rowHeight))) | Out-Null
$tblMain.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, $rowHeight))) | Out-Null
$tblMain.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, $rowHeight))) | Out-Null
$tblMain.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, $rowHeight))) | Out-Null
$tblMain.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, $rowHeight))) | Out-Null
$tblMain.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, $rowHeight))) | Out-Null
$tblMain.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, $rowHeight))) | Out-Null
$tblMain.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, $rowHeight))) | Out-Null
$tblMain.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 20))) | Out-Null
$tblMain.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,20))) | Out-Null
$tblMain.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,20))) | Out-Null
$tblMain.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,50))) | Out-Null
$tblMain.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,50))) | Out-Null
$tblMain.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,50))) | Out-Null
$tblMain.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,8))) | Out-Null
#endregion

#region Output
$global:rtbOutput = New-Object System.Windows.Forms.RichTextBox -Property @{
    dock = "Fill"
    Font = $global:font4 #//-Calibri is so much nicer
    BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    #ForeColor = [System.Drawing.Color]::FromArgb(250,250,250) #//-Dark Theme output!
    BackColor = [System.Drawing.Color]::FromArgb(20,20,20) #//-Dark Theme output! (eventually)
    #ReadOnly = $true #//-Disable editing the output box
    MultiLine = $true #//-Not sure if we need this or not
}
#endregion

#region Template
$tblTemplate = New-Object System.Windows.Forms.TableLayoutPanel -Property @{
    ColumnCount = 3
    RowCount = 2
    Dock = "Fill"
}
$tblTemplate.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40))) | Out-Null
$tblTemplate.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60))) | Out-Null
$tblTemplate.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,86))) | Out-Null
$tblTemplate.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,7))) | Out-Null
$tblTemplate.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,7))) | Out-Null
$tblMain.Controls.Add($tblTemplate,0,0)
$tblMain.SetColumnSpan($tblTemplate,3)
$lblTemplate = New-Object System.Windows.Forms.Label -Property @{
  text = "Course Settings:"
  dock = "Fill"
  Anchor = "none"
  Font = $global:font1
  TextAlign = "BottomCenter"
}
$cboTemplate = New-Object system.Windows.Forms.ComboBox -Property @{
  text                = "Click add ==========>"
  width               = 400
  #height              = 34
  #location            = New-Object System.Drawing.Point(9,25)
  Font                = $global:font1
  Dock = "none"
  Anchor = "none"
  AutoSize = $true
}
$btnTemplateAdd = New-Object System.Windows.Forms.Button -Property @{
  text = "+"
  #Dock = "fill"
  Dock = "none"
  Anchor = "none"
  Width = 30
  Font = $global:font2
  FlatStyle = "Flat"
  ForeColor = "Green"
  BackColor = "White"
}
$btnTemplateDel = New-Object System.Windows.Forms.Button -Property @{
  text = "-"
  #Dock = "fill"
  Dock = "none"
  Anchor = "none"
  Width = 30
  Font = $global:font2
  FlatStyle = "Flat"
  ForeColor = "Red"
  BackColor = "White"
}
#endregion
#region Dates
$tblDates = New-Object System.Windows.Forms.TableLayoutPanel -Property @{
  ColumnCount = 2
  RowCount = 2
  Dock = "Fill"
}

$tblDates.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40))) | Out-Null
$tblDates.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60))) | Out-Null
$tblDates.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,50))) | Out-Null
$tblDates.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,50))) | Out-Null
$tblMain.Controls.Add($tblDates,3,0)
$tblMain.SetColumnSpan($tblDates,2)
$lblStartDate = New-Object System.Windows.Forms.Label -Property @{
  text = "Class Start Date:"
  Dock = "fill"
  Anchor = "non"
  Font = $global:font1
  TextAlign = "BottomCenter"
}
$global:StartDate = New-Object System.Windows.Forms.DateTimePicker -Property @{
  Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
  CustomFormat = "MM/dd/yyy"
  Dock = "fill"
  Anchor = "none"
  Font = $global:font1
}

$lblGradeDate = New-Object System.Windows.Forms.Label -Property @{
  text = "Class Graduation Date:"
  Dock = "fill"
  Anchor = "none"
  Font = $global:font1
  TextAlign = "BottomCenter"
}

$global:EndDate = New-Object System.Windows.Forms.DateTimePicker -Property @{
  Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
  CustomFormat = "MM/dd/yyy"
  Dock = "fill"
  Anchor = "none"
  Font = $global:font1
}

#endregion

#//-Template Selection
$lblTemp = New-Object System.Windows.Forms.Label -Property @{
  text = "4419 Template:"
  Dock = "fill"
  #Anchor = "none"
  Font = $global:font1
  TextAlign = "MiddleRight"
}

$global:txtTemp = New-Object System.Windows.Forms.TextBox -Property @{
  Dock = "fill"
  Anchor = "left"
  Font = $global:font0
  TextAlign = "Left"
  width               = 1100
  #AutoSize = $true
}
$btnTemp = New-Object System.Windows.Forms.Button -Property @{
  #Dock = "fill"
  Dock = "None"
  Anchor = "left"
  Text = "..."
  Font = $global:font2
  FlatStyle = "Flat"
  ForeColor = "Black"
  BackColor = "White"
}
$FDTemp = New-Object System.Windows.Forms.OpenFileDialog -Property @{
  #InitialDirectory = [Environment]::GetFolderPath('Desktop')
  Filter = '4419 Template (*.pdf)|*.pdf'
  Title = "Select 4419 Template..."
}

#region Inputs
#//-Roster Selection
$lblRoster = New-Object System.Windows.Forms.Label -Property @{
  text = "Roster File:"
  Dock = "fill"
  #Anchor = "none"
  Font = $global:font1
  TextAlign = "MiddleRight"
}
$global:txtRoster = New-Object System.Windows.Forms.TextBox -Property @{
  Dock = "fill"
  Anchor = "left"
  Font = $global:font0
  TextAlign = "Left"
  width               = 1100
  #AutoSize = $true
}
$btnRoster = New-Object System.Windows.Forms.Button -Property @{
  #Dock = "fill"
  Dock = "None"
  Anchor = "left"
  Text = "..."
  Font = $global:font2
  FlatStyle = "Flat"
  ForeColor = "Black"
  BackColor = "White"
}
$FDRoster = New-Object System.Windows.Forms.OpenFileDialog -Property @{
  #InitialDirectory = [Environment]::GetFolderPath('Desktop')
  Filter = 'Roster (*.xlsx)|*.xlsx'
  Title = "Select Class Roster..."
}

#//-Gradebook Selection
$lblGradeBook = New-Object System.Windows.Forms.Label -Property @{
  Text = "Grade Book:"
  Dock = "fill"
  #Anchor = "none"
  Font = $global:font1
  TextAlign = "MiddleRight"
}
$global:txtGradeBook = New-Object System.Windows.Forms.TextBox -Property @{
  Dock = "fill"
  Anchor = "left"
  Font = $global:font0
  TextAlign = "Left"
  width               = 1100
  #AutoSize = $true
}
$btnGradeBook = New-Object System.Windows.Forms.Button -Property @{
  #Dock = "fill"
  Dock = "None"
  Anchor = "left"
  Text = "..."
  Font = $global:font2
  FlatStyle = "Flat"
  ForeColor = "Black"
  BackColor = "White"
}
$FDGradeBook = New-Object System.Windows.Forms.OpenFileDialog -Property @{
  #InitialDirectory = [Environment]::GetFolderPath('Desktop')
  Filter = 'GradeBook (*.xlsx)|*.xlsx'
  Title = "Select GradeBook..."
}

#//-Save Dir Selection
$lblSaveDir = New-Object System.Windows.Forms.Label -Property @{
  text = "Save Dir:"
  Dock = "fill"
  #Anchor = "none"
  Font = $global:font1
  TextAlign = "MiddleRight"
}
$global:txtSaveDir = New-Object System.Windows.Forms.TextBox -Property @{
  Dock = "fill"
  Anchor = "left"
  Font = $global:font0
  TextAlign = "Left"
  width               = 1100
  #AutoSize = $true
}
$btnSaveDir = New-Object System.Windows.Forms.Button -Property @{
  #Dock = "fill"
  Dock = "None"
  Anchor = "left"
  Text = "..."
  Font = $global:font2
  FlatStyle = "Flat"
  ForeColor = "Black"
  BackColor = "White"
}
$FDSaveDir = New-Object System.Windows.Forms.FolderBrowserDialog


#//-Instructor Title
$lblInstructor = New-Object System.Windows.Forms.Label -Property @{
  text = "Instructor:"
  Dock = "fill"
  #Anchor = "none"
  Font = $global:font1
  TextAlign = "MiddleRight"
}
$global:txtInstructor = New-Object System.Windows.Forms.TextBox -Property @{
  Dock = "fill"
  Anchor = "left"
  Font = $global:font0
  TextAlign = "Left"
  width               = 1100
  #AutoSize = $true
}

#//-Student Role Title
$lblRole = New-Object System.Windows.Forms.Label -Property @{
  text = "Student Role:"
  Dock = "fill"
  #Anchor = "none"
  Font = $global:font1
  TextAlign = "MiddleRight"
}
$global:txtRole = New-Object System.Windows.Forms.TextBox -Property @{
  Dock = "fill"
  Anchor = "left"
  Font = $global:font0
  TextAlign = "Left"
  width               = 1100
  #AutoSize = $true
}


#[String]$global:strCourseDescription
#[String]$global:strDemonstrated
#[String]$global:strAcademicExcellence
#[String]$global:strDistinguishedGraduate
#[String]$global:strOutstandingContributor

#//-Course Description String
$lblStrCourseDescription = New-Object System.Windows.Forms.Label -Property @{
  text = "Course Description:"
  Dock = "fill"
  #Anchor = "none"
  Font = $global:font1
  TextAlign = "MiddleRight"
}
$global:txtStrCourseDescription = New-Object System.Windows.Forms.TextBox -Property @{
  Dock = "fill"
  Anchor = "left"
  Font = $global:font0
  TextAlign = "Left"
  width               = 1100
  #AutoSize = $true
}

#//-Demonstrated String
$lblStrDemonstrated = New-Object System.Windows.Forms.Label -Property @{
  text = "Demonstrated:"
  Dock = "fill"
  #Anchor = "none"
  Font = $global:font1
  TextAlign = "MiddleRight"
}
$global:txtStrDemonstrated = New-Object System.Windows.Forms.TextBox -Property @{
  Dock = "fill"
  Anchor = "left"
  Font = $global:font0
  TextAlign = "Left"
  width               = 1100
  #AutoSize = $true
}

#//-AcademicExcellence String
$lblStrAcademicExcellence = New-Object System.Windows.Forms.Label -Property @{
  text = "Academic Excellence:"
  Dock = "fill"
  #Anchor = "none"
  Font = $global:font1
  TextAlign = "MiddleRight"
}
$global:txtStrAcademicExcellence = New-Object System.Windows.Forms.TextBox -Property @{
  Dock = "fill"
  Anchor = "left"
  Font = $global:font0
  TextAlign = "Left"
  width               = 1100
  #AutoSize = $true
}

#//-DistinguishedGraduate String
$lblStrDistinguishedGraduate = New-Object System.Windows.Forms.Label -Property @{
  text = "Distinguished Graduate:"
  Dock = "fill"
  #Anchor = "none"
  Font = $global:font1
  TextAlign = "MiddleRight"
}
$global:txtStrDistinguishedGraduate = New-Object System.Windows.Forms.TextBox -Property @{
  Dock = "fill"
  Anchor = "left"
  Font = $global:font0
  TextAlign = "Left"
  width               = 1100
  #AutoSize = $true
}

#//-OutstandingContributor String
$lblStrOutstandingContributor = New-Object System.Windows.Forms.Label -Property @{
  text = "Outstanding Contributor:"
  Dock = "fill"
  #Anchor = "none"
  Font = $global:font1
  TextAlign = "MiddleRight"
}
$global:txtStrOutstandingContributor = New-Object System.Windows.Forms.TextBox -Property @{
  Dock = "fill"
  Anchor = "left"
  Font = $global:font0
  TextAlign = "Left"
  width               = 1100
  #AutoSize = $true
}

#endregion
#region Progress
$global:progress = New-Object System.Windows.Forms.ProgressBar -Property @{
dock = "fill"
Style = "Continuous"
Visible = $false
Font = $global:font1
}
$global:Font = $global:font3
$global:brush1 = New-Object system.Drawing.SolidBrush([System.Drawing.Color]::Black)
#$global:PBCG = $global:progress.CreateGraphics()
[System.Drawing.Graphics]$global:PBCG = $global:progress.CreateGraphics()
#endregion
#region Buttons
$tblButtons = New-Object System.Windows.Forms.TableLayoutPanel -Property @{
  ColumnCount = 4
  RowCount = 1
  Dock = "Fill"
}

$tblButtons.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,30))) | Out-Null
$tblButtons.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,50))) | Out-Null
$tblButtons.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,10))) | Out-Null
$tblButtons.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,50))) | Out-Null

#//-Button to Process Data
$bntProcess = New-Object System.Windows.Forms.Button -Property @{
  Text = "Process"
  #Dock = "fill"
  Dock = "None"
  Anchor = "left"
  Font = $global:font3
  FlatStyle = "Flat"
  ForeColor = "Black"
  BackColor = "White"
}
#//-Button to Close App
$bntClose = New-Object System.Windows.Forms.Button -Property @{
  Text = "Close"
  #Dock = "fill"
  Dock = "None"
  Anchor = "left"
  Font = $global:font3
  FlatStyle = "Flat"
  ForeColor = "Black"
  BackColor = "White"
}

#endregion
#region Add Controls

#//-Added in Trim to help Margin out the App
$frmMain.controls.AddRange(@($tblTrim))
$tblTrim.Controls.Add($tblMain,1,1)

#//-Course selection controls
$tblTemplate.Controls.Add($cboTemplate,0,1) #Column,Row
$tblTemplate.Controls.Add($lblTemplate,0,0) #Column,Row

$tblTemplate.Controls.Add($btnTemplateAdd,1,1) #Column,Row
$tblTemplate.Controls.Add($btnTemplateDel,2,1) #Column,Row

#//-Class Start Date
$tblDates.Controls.Add($lblStartDate,0,0)
$tblDates.Controls.Add($global:StartDate,0,1)

#//-Class Graduation Date
$tblDates.Controls.Add($lblGradeDate,1,0)
$tblDates.Controls.Add($global:EndDate,1,1)

$tblMain.Dock = "Fill"

$i = 0

#//-Template file selection
$i++
$tblMain.Controls.Add($lblTemp,0,$i)
$tblMain.Controls.Add($global:txtTemp,1,$i)
$tblmain.SetColumnSpan($global:txtTemp,4)
$tblMain.Controls.Add($btnTemp,5,$i)

#//-Roster file selection
$i++
$tblMain.Controls.Add($lblRoster,0,$i)
$tblMain.Controls.Add($global:txtRoster,1,$i)
$tblmain.SetColumnSpan($global:txtRoster,4)
$tblMain.Controls.Add($btnRoster,5,$i)

#//-Gradebook file selection
$i++
$tblMain.Controls.Add($lblGradeBook,0,$i)
$tblMain.Controls.Add($global:txtGradeBook,1,$i)
$tblmain.SetColumnSpan($global:txtGradeBook,4)
$tblMain.Controls.Add($btnGradeBook,5,$i)

#//-Save directory selection
$i++
$tblMain.Controls.Add($lblSaveDir,0,$i)
$tblMain.Controls.Add($global:txtSaveDir,1,$i)
$tblmain.SetColumnSpan($global:txtSaveDir,4)
$tblMain.Controls.Add($btnSaveDir,5,$i)

#//-Instructor data field
$i++
$tblMain.Controls.Add($lblInstructor,0,$i)
$tblMain.Controls.Add($global:txtInstructor,1,$i)
$tblmain.SetColumnSpan($global:txtInstructor,4)

#//-Student Role / Crew Position data field
$i++
$tblMain.Controls.Add($lblRole,0,$i)
$tblMain.Controls.Add($global:txtRole,1,$i)
$tblmain.SetColumnSpan($global:txtRole,4)


#//-Course Description
$i++
$tblMain.Controls.Add($lblStrCourseDescription,0,$i)
$tblMain.Controls.Add($global:txtStrCourseDescription,1,$i)
$tblmain.SetColumnSpan($global:txtStrCourseDescription,4)

#//-Demonstrated
$i++
$tblMain.Controls.Add($lblStrDemonstrated ,0,$i)
$tblMain.Controls.Add($global:txtStrDemonstrated ,1,$i)
$tblmain.SetColumnSpan($global:txtStrDemonstrated ,4)

#//-AcademicExcellence
$i++
$tblMain.Controls.Add($lblStrAcademicExcellence,0,$i)
$tblMain.Controls.Add($global:txtStrAcademicExcellence,1,$i)
$tblmain.SetColumnSpan($global:txtStrAcademicExcellence,4)

#//-DistinguishedGraduate
$i++
$tblMain.Controls.Add($lblStrDistinguishedGraduate,0,$i)
$tblMain.Controls.Add($global:txtStrDistinguishedGraduate,1,$i)
$tblmain.SetColumnSpan($global:txtStrDistinguishedGraduate,4)

#//-OutstandingContributor
$i++
$tblMain.Controls.Add($lblStrOutstandingContributor,0,$i)
$tblMain.Controls.Add($global:txtStrOutstandingContributor,1,$i)
$tblmain.SetColumnSpan($global:txtStrOutstandingContributor,4)



#//-Progress bar
$i++
$tblMain.Controls.Add($global:progress,0,$i)
$tblMain.SetColumnSpan($global:progress,4)
$tblMain.Controls.Add($tblButtons,4,$i)

#//-Output display pane
$i++
$tblMain.Controls.Add($global:rtbOutput,0,$i)
$tblMain.SetColumnSpan($global:rtbOutput,6)
#$tblMain.SetRowSpan($global:rtbOutput,2)

#//-Process Button
$tblButtons.Controls.Add($bntProcess,1,0)
#//-Close Button
$tblButtons.Controls.Add($bntClose,3,0)


#$frmMain.CancelButton = $bntClose
#endregion
#endregin
#region Event Listeners

$bntClose.Add_Click({SaveSettings;$frmMain.Close()})
$frmMain.Add_Load({ LoadSettings })

$btnTemplateDel.add_click({
  DeleteTemplate -templateName $cboTemplate.SelectedItem
})
$bntProcess.add_click({
  SaveSettings
  if ($global:txtGradeBook.Text -ne "" -and $global:txtRoster.Text -ne "" -and $global:txtTemp.Text -ne "" -and $global:txtSaveDir.Text -ne "" -and $global:txtInstructor.Text -ne "" -and $global:txtRole.Text -ne ""){
    UpdatePDF
  }else{
    [System.Windows.Forms.MessageBox]::Show("Please Complete the Required Boxes")
  }
})
$global:txtGradeBook.Add_TextChanged({ $global:currentTemplate.GradeBookFile = $global:txtGradeBook.Text })
$global:txtTemp.Add_TextChanged({ $global:currentTemplate.TemplateFile = $global:txtTemp.Text })
$global:txtRoster.Add_TextChanged({ $global:currentTemplate.RosterFile = $global:txtRoster.Text })
$global:txtSaveDir.Add_TextChanged({ $global:currentTemplate.SaveDir = $global:txtSaveDir.Text })
$global:txtInstructor.Add_TextChanged({ $global:currentTemplate.Instructor = $global:txtInstructor.Text })
$global:txtRole.Add_TextChanged({ $global:currentTemplate.Role = $global:txtRole.Text })

$global:txtStrCourseDescription.Add_TextChanged({ $global:currentTemplate.StrCourseDescription = $global:txtStrCourseDescription.Text })
$global:txtStrDemonstrated.Add_TextChanged({ $global:currentTemplate.StrDemonstrated = $global:txtStrDemonstrated.Text })
$global:txtStrAcademicExcellence.Add_TextChanged({ $global:currentTemplate.StrAcademicExcellence = $global:txtStrAcademicExcellence.Text })
$global:txtStrDistinguishedGraduate.Add_TextChanged({ $global:currentTemplate.StrDistinguishedGraduate = $global:txtStrDistinguishedGraduate.Text })
$global:txtStrOutstandingContributor.Add_TextChanged({ $global:currentTemplate.StrOutstandingContributor = $global:txtStrOutstandingContributor.Text })


$global:StartDate.Add_TextChanged({$global:currentTemplate.startDate = $global:StartDate.Text})
$global:EndDate.Add_TextChanged({$global:currentTemplate.EndDate = $global:EndDate.Text})


$cboTemplate.Add_MouseEnter({
  if($cboTemplate.Items.Count -gt 0){
    SaveSettings
   }
})
$cboTemplate.Add_SelectedIndexChanged({
  $global:currentTemplate = $global:templates | Where-Object {$_.templateName -eq $cboTemplate.SelectedItem}
  $global:txtTemp.Text = $global:currentTemplate.TemplateFile
  $global:txtRoster.Text = $global:currentTemplate.RosterFile
  $global:txtGradeBook.Text = $global:currentTemplate.GradeBookFile
  $global:txtSaveDir.Text = $global:currentTemplate.SaveDir
  $global:txtInstructor.Text = $global:currentTemplate.Instructor
  $global:txtRole.Text = $global:currentTemplate.Role
  $global:EndDate.Text = $global:currentTemplate.EndDate
  $global:StartDate.Text = $global:currentTemplate.startDate
  #//-Need to add the customization through GUI for these string values
  $global:txtStrCourseDescription.Text = $global:currentTemplate.StrCourseDescription
  $global:txtStrDemonstrated.Text = $global:currentTemplate.StrDemonstrated
  $global:txtStrAcademicExcellence.Text = $global:currentTemplate.StrAcademicExcellence
  $global:txtStrDistinguishedGraduate.Text = $global:currentTemplate.StrDistinguishedGraduate
  $global:txtStrOutstandingContributor.Text = $global:currentTemplate.StrOutstandingContributor

})


$btnTemplateAdd.add_click({
  add-Type -AssemblyName Microsoft.VisualBasic
  $title = 'Demographics'
  $msg   = 'Enter New Template Name:'
  $text = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
  #//-Needs an empty/cancel string catch for blank names
  if($text.trim() -eq "") { return }

  $cboTemplate.Items.Add($text)
  $global:currentTemplate = New-Object Template
  $global:currentTemplate.TemplateName = $text

  #//-Need to add in customization for these
  $global:txtStrCourseDescription.Text      = $global:strDEFAULTCourseDescription
  $global:txtStrDemonstrated.Text           = $global:strDEFAULTDemonstrated
  $global:txtStrAcademicExcellence.Text     = $global:strDEFAULTAcademicExcellence
  $global:txtStrDistinguishedGraduate.Text  = $global:strDEFAULTDistinguishedGraduate
  $global:txtStrOutstandingContributor.Text = $global:strDEFAULTOutstandingContributor

  #Append to current array of Templates
  $global:templates += $global:currentTemplate
  $cboTemplate.SelectedIndex = $cboTemplate.FindStringExact($text)
})
$btnGradeBook.add_click({
  if ($FDGradeBook.ShowDialog() -eq "OK"){
    $global:txtGradeBook.Text = $FDGradeBook.Filename
  }
  })
$btnRoster.add_click({
  if ($FDRoster.ShowDialog() -eq "OK") {
    $global:txtRoster.Text = $FDRoster.Filename
  } })
$btnTemp.add_click({
  if ($FDTemp.ShowDialog() -eq "OK"){
    $global:txtTemp.Text = $FDTemp.Filename
  }})
$btnSaveDir.add_click({
  if ($FDSaveDir.ShowDialog() -eq "OK"){
    $global:txtSaveDir.Text = $FDSaveDir.SelectedPath
  } })
#endregion


#endregion

$tblTemplate.CellBorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$tblMain.CellBorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$tblDates.CellBorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$tblButtons.CellBorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

$frmMain.ShowDialog() | Out-Null
$frmMain.Focus() | Out-Null
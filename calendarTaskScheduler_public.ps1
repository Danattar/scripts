Function Get-Tasks($path, $schedule)
##fonction pour recuperer toutes les tâches planifiés du serveur specifie dans la variable schedule. $path specifie le repertoire de tâches a selectionner.
{
	$out = @()
	$schedule.GetFolder($path).GetTasks(0) | % {
		$xml = [xml]$_.xml
        #construit un object avec les parametres des taches planifiees
        $out += New-Object psobject -Property @{
		

		    "Name" = $_.Name
			"Actions" = ($xml.Task.Actions.Exec | % { "$($_.Command) $($_.Arguments)" }) -join "`n"
			"Repetition.Interval" = $(If($xml.task.triggers.CalendarTrigger.Repetition){$xml.task.triggers.CalendarTrigger.Repetition.Interval})
            "Repetition.Duration" = $(IF($xml.task.triggers.CalendarTrigger.Repetition){$xml.task.triggers.CalendarTrigger.Repetition.Duration})
            "StartBoundary" =$(If($xml.task.triggers.CalendarTrigger.startBoundary){[DateTime]$xml.task.triggers.CalendarTrigger.startBoundary}else{Get-Date})
            "EndBoundary" = $(If($xml.task.triggers.CalendarTrigger.endBoundary){[DateTime]$xml.task.triggers.CalendarTrigger.endBoundary}else{$defaultEndBoundary})
            "ScheduleByDay" = $(If($xml.task.triggers.CalendarTrigger.ScheduleByDay){$xml.task.triggers.CalendarTrigger.ScheduleByDay.DaysInterval}else{""})
            "ScheduleByWeek.DaysOfWeek" = $(If($xml.task.triggers.CalendarTrigger.ScheduleByWeek){ForEach($task in($xml.task.triggers.CalendarTrigger.ScheduleByWeek.DaysOfWeek| gm | Where{$_.membertype -eq "Property"})){$task.name}}else{""})
            "ScheduleByWeek.WeeksInterval" = $(If($xml.task.triggers.CalendarTrigger.ScheduleByWeek.WeeksInterval){$xml.task.triggers.CalendarTrigger.ScheduleByWeek.WeeksInterval})
            "ScheduleByMonth.DaysOfMonth" = $(If($xml.task.triggers.CalendarTrigger.ScheduleByMonth){ForEach($task in($xml.task.triggers.CalendarTrigger.ScheduleByMonth.DaysOfMonth| gm | Where{$_.membertype -eq "Property"})){$xml.task.triggers.CalendarTrigger.ScheduleByMonth.DaysOfMonth.$($task.name)}}<#else{""}#>)
            "ScheduleByMonth.Months" = $(If($xml.task.triggers.CalendarTrigger.ScheduleByMonth){ForEach($task in($xml.task.triggers.CalendarTrigger.ScheduleByMonth.Months| gm | Where{$_.membertype -eq "Property"})){$task.name}}<#else{""}#>)
            "ScheduleByMonthDOW.Weeks" = $(If($xml.task.triggers.CalendarTrigger.ScheduleByMonthDayOfWeek.Weeks){ForEach($task in ($xml.task.triggers.CalendarTrigger.scheduleByMonthDayOfWeek.Weeks | gm | Where{$_.membertype -eq "Property"})){$xml.task.triggers.CalendarTrigger.scheduleByMonthDayOfWeek.Weeks.$($task.name)}}else{})
            "ScheduleByMonthDOW.DOW" = $(If($xml.task.triggers.CalendarTrigger.ScheduleByMonthDayOfWeek){ForEach($task in ($xml.task.triggers.CalendarTrigger.scheduleByMonthDayOfWeek.DaysOfWeek | gm | Where{$_.membertype -eq "Property"})){$task.name}}else{})
            "ScheduleByMonthDOW.Months" = $(If($xml.task.triggers.CalendarTrigger.ScheduleByMonthDayOfWeek){ForEach($task in($xml.task.triggers.CalendarTrigger.ScheduleByMonthDayOfWeek.Months| gm | Where{$_.membertype -eq "Property"})){$task.name}}<#else{""}#>)
            "Author" = $xml.task.principals.Principal.UserID
			"Description" = $xml.task.registrationInfo.Description
                    
					
        }
               
	}
	If(!$RootOnly)
	{
		$schedule.GetFolder($path).GetFolders(0) | % {
			$out += get-Tasks($_.Path)
		}
	}
	$out
}
Function Get-ScheduledTasks
{
	Param
	(
		[Alias("Computer","localhost")]
		[Parameter(Position=1,ValuefromPipeline=$true,ValuefromPipelineByPropertyName=$true)]
		[string[]]$Name = $env:COMPUTERNAME
        ,[switch]$RootOnly = $false
	)
	Begin
	{
		$tasks = @()
		$schedule = New-Object -ComObject "Schedule.Service"
        $schedule.connect($Computer)
        $path = $schedule.getfolder("\folder").GetFolders(0)
  	}
	Process
	{
        ForEach($Computer in $Name)
		        {
			        If(Test-Connection $computer -count 1 -quiet)
			        {
				        #$schedule.connect($Computer)
                        foreach($paths in $path){
				            $tasks += Get-Tasks $paths.Path $schedule
                        }
			        }
			        Else
			        {
				        Write-Error "Cannot connect to $Computer. Please check it's network connectivity."
				        Break
			        }
			        $tasks
		        }
	        }
	        End
	        {
		        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($schedule) | Out-Null
		        Remove-Variable schedule
	        }
}
function timetable($repetition){

#function to convert scheduler interval and duration format to minutes.
    switch($repetition.'Interval'){
        'PT5M'{$repetition.'Interval' = 5}
        'PT10M'{$repetition.'Interval' = 10}
        'PT15M'{($repetition.'Interval') = 15}
        'PT30M'{$repetition.'Interval' = 30}
        'PT1H'{$repetition.'Interval' = 60}
        default{$repetition.'Interval' = 5}
    }
    switch($repetition.'Duration'){
        'PT15M'{$repetition.'Duration' = 15}
        'PT30M'{$repetition.'Duration' = 30}
        'PT1H'{$repetition.'Duration' = 60}
        'PT12H'{$repetition.'Duration' = 720}
        'P1D'{$repetition.'Duration' = 1440}
        default{$repetition.'Duration' = 5}
    }
    $repetition.'nbOfRepetition'= ($repetition.'Duration')/($repetition.Interval)
    $repetition

}

function Get-Task-Runtime($task){
    $stockStarts =Get-WinEvent -FilterHashtable @{logname='Microsoft-Windows-TaskScheduler/Operational'; ID=200; level=4}#à optimiser
    $stockStops =Get-WinEvent -FilterHashtable @{logname='Microsoft-Windows-TaskScheduler/Operational'; ID=201; level=4}#à optimiser
    $times= @()
    $i =0
    foreach ($startItem in $stockStarts){
        $mess = $startItem.message
        if ($mess -match $task.Name){
            #$instance = match '"{[a-zA-Z_0-9]*}"'
            $pos = $mess.IndexOf("{")
            $newmess = $mess.Substring($pos+1)
            $pos = $newmess.IndexOf("}")
            $instance = $newmess.SubString(0,$pos)

            foreach($stopItem in $stockStops){
                $mess2 = $stopItem.message
                if($mess2 -match $instance){
                $endTime = $stopItem.TimeCreated
                $startTime = $startItem.TimeCreated
                $times += ($endTime - $startTime)
                $i = $i+1
                break                                         #fonctionnel sale...
                }
            }
        }
        if ($i -eq 5){
        break
        }
    }
    $finalTime = 0
    foreach ($time in $times){
        $finalTime = $finalTime + $time
    }
    if($finalTime.Milliseconds -gt 500){
        $seconds = ($finalTime.seconds +1)
    }
    else{
        $seconds = $finalTime.seconds
    }
    if($finalTime.days -ne 0){
        $seconds = $seconds + ($finalTime.Days * 86400)
    }
    if($finalTime.Hours -ne 0){
        $seconds = $seconds + ($finalTime.Hours * 3600)
    }
    if($finaltime.Minutes -ne 0){
        $seconds = $seconds +($finalTime.Minutes * 60)
    }
    if($times.Count -gt 0){
        $finalTime = ($seconds/$times.Count)
    }
    else{
        $finalTime = 900
    }
    if($finalTime -lt 900){
        $finalTime = 900
    }
    $task | Add-Member TaskTime $finalTime
    $task
}
function FillCalendar($tasks){

    BEGIN { 
        #creer les rendez-vous dans le calendrier pour la $tasks
   
    }

    PROCESS { 
        #initialize connexion to exchange
        Add-Type -Path "C:\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"
        $version = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
        $service = new-object Microsoft.Exchange.WebServices.Data.ExchangeService($version)  
        $uri=[system.URI] "https://exchange.server.com"
        $service.Url = $uri
        #Initialize variables
   
        if($tasks.'ScheduleByWeek.DaysOfWeek')
        {
            $days= @()
            $days=$tasks.'ScheduleByWeek.DaysOfWeek'
            $frequency =$tasks.'ScheduleByWeek.WeeksInterval'
            foreach ($day in $days)
            {
                $starts = $tasks.StartBoundary

                for($j=1; $j -le 7; $j++)
                {        
                    if($starts.AddDays($j).DayOfWeek -eq $day)
                    {
                        $starts = $starts.AddDays($j)
                        break
                    }
                }
                #create new appointment
                if($tasks.'Repetition.Interval'){
                    if(-not $tasks.'Repetition.Duration'){  #if repetition duration set to indefinetely
                        
                        
                        $body = "<!doctype HTML><head><meta charset='UTF-8'></head><body>Scheduled Task "+$tasks.name + " from server " +$env:COMPUTERNAME+ " was set with a repetition duration set to indefinitely.</body>"
                        Send-MailMessage -To 'lorem@dolor.sit' -Subject 'A scheduled task was wrong parametred' -Body $body -BodyAsHtml -SmtpServer 'mail.ipsum.com' -From 'lorem.ipsum@dolor.sit'
                    }
                    else{
                        $repetition = New-Object psobject -Property @{
                        "Interval"= $tasks.'Repetition.Interval'
                        "Duration"= $tasks.'Repetition.Duration'
                        "nbOfRepetition" = 0
                        }
                         $repetition = timetable($repetition)

                        for($i = 0 ; $i -lt $repetition.'nbOfRepetition' ;$i++){
                                $newApt = new-object Microsoft.Exchange.WebServices.Data.Appointment($service)
                                #$newApt.Start = [DATETIME]::Now
                                $start = ($starts).AddMinutes($repetition.'Interval'*$i)
                                $newApt.Start =  $start
                                $newApt.End = ($start).addSeconds($tasks.'TaskTime')
                                $newApt.Subject = $tasks.'Name'
                                $newApt.Body = $tasks.'Description'
                                $DayOfTheWeek = New-Object Microsoft.Exchange.WebServices.Data.DayOfTheWeek[] 1
                                $DayOfTheWeek[0] = [Microsoft.Exchange.WebServices.Data.DayOfTheWeek]::$day
                                $newApt.Recurrence = New-Object Microsoft.Exchange.WebServices.Data.Recurrence+WeeklyPattern($start, $frequency, $DayOfTheWeek);
                                $newApt.Recurrence.StartDate =$start
                                if($tasks.'EndBoundary'){
                                    $newApt.Recurrence.EndDate = $tasks.'EndBoundary'
                                }else{
                                    $newApt.Recurrence.EndDate = $start.addYears(2)
                                }
                                $newApt.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToNone)
                        }
                    }
                }
                else{
                    $newApt = new-object Microsoft.Exchange.WebServices.Data.Appointment($service)
                    #$newApt.Start = [DATETIME]::Now
                    $newApt.Start =  $starts
                    $newApt.End = ($starts).addSeconds($tasks.'TaskTime')
                    $newApt.Subject = $tasks.'Name'
                    $newApt.Body = $tasks.'Description'
                    $DayOfTheWeek = New-Object Microsoft.Exchange.WebServices.Data.DayOfTheWeek[] 1
                    $DayOfTheWeek[0] = [Microsoft.Exchange.WebServices.Data.DayOfTheWeek]::$day
                    $newApt.Recurrence = New-Object Microsoft.Exchange.WebServices.Data.Recurrence+WeeklyPattern($starts, $frequency, $DayOfTheWeek);
                    $newApt.Recurrence.StartDate =$starts
                    if($tasks.'EndBoundary'){
                        $newApt.Recurrence.EndDate = $tasks.'EndBoundary'
                    }else{
                        $newApt.Recurrence.EndDate = $starts.addYears(2)
                    }
                    $newApt.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToNone)
                }
            }
        }
        elseif($tasks.'ScheduleByMonth.Months')
        {
            
            foreach($month in $tasks.'ScheduleByMonth.Months'){
                foreach($day in $tasks.'ScheduleByMonth.DaysOfMonth'){
                    
                    Switch($month)
                    {
                        'January'{$month = '01'}
                        'February'{$month = '02'}
                        'March'{$month = '03'}
                        'April'{$month = '04'}
                        'May'{$month = '05'}
                        'June'{$month = '06'}
                        'July'{$month = '07'}
                        'August'{$month = '08'}
                        'September'{$month = '09'}
                        'October'{$month = '10'}
                        'November'{$month = '11'}
                        'December'{$month = '12'}
                        default{}
                    }

                    if(($day -as [int]) -lt 10)
                    {
                        $date = $month+'/'+ '0'+ $day+'/'+(get-date).Year
                    }
                    else
                    {
                        $date = $month+'/'+$day+'/'+(get-date).Year
                    }

                    if($tasks.'Repetition.Interval'){
                        if(-not $tasks.'Repetition.Duration'){  #if repetition duration set to indefinetely
                        
                            
                            $body = "<!doctype HTML><head><meta charset='UTF-8'></head><body>Scheduled Task "+$tasks.name + " from server " +$env:COMPUTERNAME+ " was set with a repetition duration set to indefinitely.</body>"
                            Send-MailMessage -To 'lorem@dolor.sit' -Subject 'A scheduled task was wrong parametred' -Body $body -BodyAsHtml -SmtpServer 'mail.fr.ch' -From 'ipsum@dolor.sit'
                        }
                        else{
                            $repetition = New-Object psobject -Property @{
                            "Interval"= $tasks.'Repetition.Interval'
                            "Duration"= $tasks.'Repetition.Duration'
                            "nbOfRepetition" = 0
                            }
                             $repetition = timetable($repetition)

                            for($i = 0 ; $i -lt $repetition.'nbOfRepetition' ;$i++){
                                $frequency = 12
                                #create new appointment
                                $newApt = new-object Microsoft.Exchange.WebServices.Data.Appointment($service)
                                #$newApt.Start = [DATETIME]::Now
                                $date =[datetime]$date
                                $start = ($date).AddMinutes($repetition.'Interval'*$i)
                                $newApt.Start = $start
                                $newApt.End = ($start).addSeconds($tasks.'TaskTime')
                                $newApt.Subject = $tasks.'Name'
                                $newApt.Body = $tasks.'Description'
                                $newApt.Recurrence = New-Object Microsoft.Exchange.WebServices.Data.Recurrence+MonthlyPattern($start, $frequency, $day); 
                                $newApt.Recurrence.StartDate =$start
                                if($tasks.'EndBoundary'){
                                    $newApt.Recurrence.EndDate = $tasks.'EndBoundary'
                                }else{
                                    $newApt.Recurrence.EndDate = $tasks.'StartBoundary'.addyears(2)
                                }
                                $newApt.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToNone)
                            }
                        }
                    }
                    else{
                        $frequency = 12
                        #create new appointment
                        $newApt = new-object Microsoft.Exchange.WebServices.Data.Appointment($service)
                        #$newApt.Start = [DATETIME]::Now
                        $newApt.Start = $date
                        $newApt.End = ($date).addSeconds($tasks.'TaskTime')
                        $newApt.Subject = $tasks.'Name'
                        $newApt.Body = $tasks.'Description'
                        $newApt.Recurrence = New-Object Microsoft.Exchange.WebServices.Data.Recurrence+MonthlyPattern($date, $frequency, $day); 
                        $newApt.Recurrence.StartDate =$date
                        if($tasks.'EndBoundary'){
                            $newApt.Recurrence.EndDate = $tasks.'EndBoundary'
                        }else{
                            $newApt.Recurrence.EndDate = $tasks.'StartBoundary'.addyears(2)
                        }
                        $newApt.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToNone)
                        $i= $i+1

                    }
                }
            }
        }
        elseif($tasks.'ScheduleByMonthDOW.Months')
        {
            $dates = @()
            foreach($month in $tasks.'ScheduleByMonthDOW.Months')
            {
                foreach($week in $tasks.'ScheduleByMonthDOW.Weeks')
                {
                    foreach($day in $tasks.'ScheduleByMonthDOW.DOW')
                    {
                       Switch($month)
                        {
                            'January'{$month = '01'}
                            'February'{$month = '02'}
                            'March'{$month = '03'}
                            'April'{$month = '04'}
                            'May'{$month = '05'}
                            'June'{$month = '06'}
                            'July'{$month = '07'}
                            'August'{$month = '08'}
                            'September'{$month = '09'}
                            'October'{$month = '10'}
                            'November'{$month = '11'}
                            'December'{$month = '12'}
                            default{}
                        }
                    
                        ###################
                        #Check si la tâche s'éffectue la dernière semaine du mois
                        ###################
                        If($week -eq 'Last')
                        {
                            $lastDay = 0
                            switch($month)
                            {

                                '01'{$lastDay = 31}
                                '03'{$lastDay = 31}
                                '05'{$lastDay = 31}
                                '07'{$lastDay = 31}
                                '08'{$lastDay = 31}
                                '10'{$lastDay = 31}
                                '12'{$lastDay = 31}
                                '04'{$lastDay = 30}
                                '06'{$lastDay = 30}
                                '09'{$lastDay = 30}
                                '11'{$lastDay = 30}
                                '02'
                                    {      #Calcul année bissextile
                                        $modulo = $tasks.StartBoundar.Year
                                        $modulo2 = $tasks.StartBoundar.Year
                                        $modulo3 = $tasks.StartBoundar.Year
                                        $modulo%= 4
                                        $modulo2%= 100
                                        $modulo3%= 400
                                        if($modulo -eq 0)
                                        {
                                            if($modulo2 -ne 0 )
                                            {
                                                $lastDay = 28
                                            }
                                            elseif($modulo3 -ne 0)
                                            {
                                                $lastDay = 29
                                            }
                                        }
                                        else
                                        {
                                            $lastDay = 28
                                        }
                                    }
                            }
                       
                            $date = [DateTime]($month+'/'+$lastDay+'/'+$tasks.startBoundary.Year)
                           # $weeknbr= get-date -date $date -UFormat %W
                            for($j=0;$j -lt 2; $j++){
                                $date = $date.AddYears($j)

                                for ($i = 0; $i -lt 7; $i++){
                                    $testDate = $date.AddDays(-$i)
                                    if($testDate.DayOfWeek -eq $day)
                                    {
                                        break
                                    }
                                }
                            
                            $date = $date.AddDays(-$i)
                            if(-not ($date -in $dates)){
                            $dates+=$date  
                            }
                            }
                            
                        }
                        #################
                        #fin du check if dernière semaine du mois
                        #################
                        #Get le jour de la dernière semaine 
                        #####################################
                        
                        
                        else{
                            $date = [datetime]($month+'/'+'01'+'/'+$tasks.startBoundary.Year) #month/day/year
                            switch($week){
                                '1'{$dayToAdd = 0}
                                '2'{$dayToAdd = 7}
                                '3'{$dayToAdd = 14}
                                '4'{$dayToAdd = 21}
                            }
                            for ($i = 0; $i -lt 7; $i++){
                                $testDate = $date.AddDays($i)
                                if($testDate.DayOfWeek -eq $day)
                                {
                                    break
                                }
                            }
                            $date = $date.AddDays($i+$dayToAdd)
                            if(-not ($date -in $dates)){
                            $dates+=$date  
                            }
                        } 
                       
                   }
                }
            }
            #$frequency = 12
            #create new appointment
            foreach($date in $dates){

                    


                if($tasks.'Repetition.Interval'){
                    if(-not $tasks.'Repetition.Duration'){  #if repetition duration set to indefinetely
                        
                        
                        $body = "<!doctype HTML><head><meta charset='UTF-8'></head><body>Scheduled Task "+$tasks.name + " from server " +$env:COMPUTERNAME+ " was set with a repetition duration set to indefinitely.</body>"
                        Send-MailMessage -To 'lorem@dolor.sit' -Subject 'A scheduled task was wrong parametred' -Body $body -BodyAsHtml -SmtpServer 'mail.fr.ch' -From 'ipsum@dolor.sit'
                    }
                    else{
                        $repetition = New-Object psobject -Property @{
                        "Interval"= $tasks.'Repetition.Interval'
                        "Duration"= $tasks.'Repetition.Duration'
                        "nbOfRepetition" = 0
                        }
                         $repetition = timetable($repetition)

                            for($i = 0 ; $i -lt $repetition.'nbOfRepetition' ;$i++){
                            

                                $newApt = new-object Microsoft.Exchange.WebServices.Data.Appointment($service)

                                $start = ($date).AddMinutes($repetition.'Interval'*$i)
                                #$newApt.Start = [DATETIME]::Now
                                $newApt.Start = $start
                                $newApt.End = ($start).addSeconds($tasks.'TaskTime')
                                $newApt.Subject = $tasks.'Name'
                                $newApt.Body = $tasks.'Description'
                                $newApt.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToNone)
                            }
                        }
                    }
                    else{
                        $newApt = new-object Microsoft.Exchange.WebServices.Data.Appointment($service)
                        $newApt.Start = $date
                        $newApt.End = ($date).addSeconds($tasks.'TaskTime')
                        $newApt.Subject = $tasks.'Name'
                        $newApt.Body = $tasks.'Description'
                        $newApt.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToNone)


                    }
                 }
        }
        elseif($tasks.'ScheduleByDay')
        {
            if($tasks.'Repetition.Interval'){
                if(-not $tasks.'Repetition.Duration'){  #if repetition duration set to indefinetely
                        
                    
                    $body = "<!doctype HTML><head><meta charset='UTF-8'></head><body>Scheduled Task "+$tasks.name + " from server " +$env:COMPUTERNAME+ " was set with a repetition duration set to indefinitely.</body>"
                    Send-MailMessage -To 'lorem@dolor.sit' -Subject 'A scheduled task was wrong parametred' -Body $body -BodyAsHtml -SmtpServer 'mail.fr.ch' -From 'ipsum@dolor.sit'
                }
                else{
                    $repetition = New-Object psobject -Property @{
                    "Interval"= $tasks.'Repetition.Interval'
                    "Duration"= $tasks.'Repetition.Duration'
                    "nbOfRepetition" = 0
                    }
                        $repetition = timetable($repetition)

                        for($i = 0 ; $i -lt $repetition.'nbOfRepetition' ;$i++){
                            $date = $tasks.'StartBoundary'
                            $frequency = $tasks.'ScheduleByDay'
                            $newApt = new-object Microsoft.Exchange.WebServices.Data.Appointment($service)
                            $start = ($date).AddMinutes($repetition.'Interval'*$i)
                            $newApt.Start = $start
                            $newApt.End = ($start).addSeconds($tasks.'TaskTime')
                            $newApt.Subject = $tasks.'Name'
                            $newApt.Body = $tasks.'Description'
                            $newApt.Recurrence = New-Object Microsoft.Exchange.WebServices.Data.Recurrence+DailyPattern($date,$frequency);
                            if($tasks.'EndBoundary'){
                                $newApt.Recurrence.EndDate = $tasks.'EndBoundary'
                            }else{
                                $newApt.Recurrence.EndDate = $date.addyears(2)
                            }
                            $newApt.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToNone)
                        }
                    }
                }
                else{
                    $date = $tasks.'StartBoundary'
                    $frequency = $tasks.'ScheduleByDay'
                    $newApt = new-object Microsoft.Exchange.WebServices.Data.Appointment($service)
                    $newApt.Start = $date
                    $newApt.End = ($date).addSeconds($tasks.'TaskTime')
                    $newApt.Subject = $tasks.'Name'
                    $newApt.Body = $tasks.'Description'
                    $newApt.Recurrence = New-Object Microsoft.Exchange.WebServices.Data.Recurrence+DailyPattern($date,$frequency);
                    if($tasks.'EndBoundary'){ 
                        $newApt.Recurrence.EndDate = $tasks.'EndBoundary'
                    }else{
                        $newApt.Recurrence.EndDate = $date.addyears(2)
                    }
                    $newApt.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToNone)
            
                }
            }
        }

    END { 
     
    }
}


$tasks = Get-ScheduledTasks
foreach($task in $tasks){
    $task = Get-Task-Runtime($task)
    FillCalendar($task)

}

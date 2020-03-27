$Error.Clear()

function get-ReportFile ()
{
    param ($InputFile)

    function get-IncLocation(){
        
        param($TMPValue, $IncGroup)

        $str=$TMPValue
        
        function get-IncLocationKrasn()
        {
            param($TMPsplitValue)

            $CountWord = $TMPsplitValue.Count
            $TMPValue1 = $TMPsplitValue[$CountWord-3][0] + $TMPsplitValue[$CountWord-3].Substring(1).tolower()
            #if(!$?){$TMPValue1 = $TMPsplitValue}   ### Если не читается имя то что мне сделать?
            $TMPValue2 = $TMPsplitValue[0..($CountWord-4)]  

            return $TMPValue1, $TMPValue2
        }

        function get-IncLocationUvst()
        {
            param($TMPsplitValue)

            $CountWord = $TMPsplitValue.Count
            $TMPValue1 = $TMPsplitValue[0][0] + $TMPsplitValue[0].Substring(1).tolower()
            $TMPValue2 = ""

            return $TMPValue1, $TMPValue2
        }
        
        function get-IncLocationZsib()
        {
            param($TMPsplitValue)

            $CountWord = $TMPsplitValue.Count
            $TMPValue1 = $TMPsplitValue[0][0] + $TMPsplitValue[0].Substring(1).tolower()
            $TMPValue2 = ""

            return $TMPValue1, $TMPValue2
        }


        if($str -notlike ""){
            $startstr=$str.IndexOf('Расположение')
            $TMPValue = $str.Substring($startstr+14).split(' ')[0].replace('-',' ')
            $TMPsplitValue = $TMPValue.Split()  

            switch($IncGroup){
                'СМЕНА-СТО-КРАСН' {$TMPValue1,$TMPValue2 = get-IncLocationKrasn $TMPsplitValue}
                'СМЕНА-СТО-ЮВСТ' {$TMPValue1,$TMPValue2 = get-IncLocationUvst $TMPsplitValue}
                'СМЕНА-СТО-ЗСИБ' {$TMPValue1,$TMPValue2 = get-IncLocationZsib $TMPsplitValue}
            }
            
            #if($TMPValue1 -like $TMPsplitValue){Write-Host $TMPValue}   ### Проверка на несоответствие записи из строки №19
        }

        return $TMPValue1,$TMPValue2
    }

    function get-IncedentObjet()
    {
        param ($InputFile, $LocationGlossary)

        $ImportFile = Get-Content $InputFile

        if($ImportFile[0] -notlike 'Номер*Краткое описание*Статус*Рабочая группа*Время направления в работу*Время закрытия события в системе мониторинга*Фактическое завершение*Решение*Корневая причина*Технический*Инфраструктурный*Подробное описание*Исполнитель*Крайний срок'){
            [System.Windows.MessageBox]::Show("Ошибка формата входного файла: `n'$InputFile' `nПроверьте правильность импортированнаго файла")
        }else{
            $Property=$ImportFile[0].Split("`*")
            $Incidents=@()
            foreach($Incident in $ImportFile[1..$ImportFile.Count])
            {
                $Object=New-Object psobject
                for($i=0; $i -lt $Property.Count;$i++)
                {
                    $TMPValue=$Incident.Split("`*")[$i]
                    $Object|Add-Member -MemberType NoteProperty -Name $Property[$i] -Value $TMPValue -Force

                    #Location properties
                    if($i -eq 11){
                        [string]$TMPValue1,[string]$TMPValue2 = get-IncLocation -TMPValue $TMPValue -IncGroup $Object.'Рабочая группа'

                        if($TMPValue1 -like 'из'){Write-Host $Incident}

                        #$Object|Add-Member -MemberType NoteProperty -Name $Property[$i] -Value $TMPValue -Force       Старая Запись

                        ### Исключение для необычных имен
                        if($TMPValue1 -like "Квосточный"){$TMPValue1 = "Красноярск-Восточный"}
                        if($TMPValue1 -like "Конторадсзлобино"){$TMPValue1 = "Злобино"; $TMPValue2 = "Контора ДС"}

                        $Object|Add-Member -MemberType NoteProperty -Name 'Станция' -Value $TMPValue1 -Force
                        $Object|Add-Member -MemberType NoteProperty -Name 'Узел' -Value $TMPValue2 -Force
                    }

                    $RootCause = @("Аварийное отключение электроэнергии","Обрыв/повреждение кабеля электропитания", `
                                    "Отключение электропитания городскими электросетями", `
                                    "Плановые работы энергетиков, без оповещения ИВЦ")
                    if($Property[$i] -match "Корневая причина") 
                    {
                        if($RootCause -contains $TMPValue)
                        {
                            $Object|Add-Member -MemberType NoteProperty -Name 'Электроэнергия' -Value $true -Force
                        }else{
                            $Object|Add-Member -MemberType NoteProperty -Name 'Электроэнергия' -Value $false -Force
                        }
                    }


                    #convert String to Time type
                    $TimePropertiyes=@("Время направления в работу", "Крайний срок", `
                                        "Время закрытия события в системе мониторинга", "Фактическое завершение")
                    if($TimePropertiyes -contains $Property[$i] -and $TMPValue -notlike "")
                    {                     
                        $Object.[string]$Property[$i]= (get-date($TMPValue))
                    }
                }
                $Incidents+=$Object

                $TMPValue,$TMPValue1,$TMPValue2 = ""
            }
        }
        return $Incidents
    }


    if(!$InputFile){$InputFile =(dir "$env:USERPROFILE\Downloads\" export*.txt | Sort-Object LastWriteTime | select -Last 1).fullname}

    return (get-IncedentObjet $InputFile)
}

function add-StationsWhereDublicateInc()
{
    param($Incidents)

    $GroupDubleIncidents = $Incidents | Where {$_.'Решение' -match 'Дублирование инцидента' -and $_.'Рабочая группа' -like "СМЕНА-СТО-КРАСН"} `
        | Group-Object 'Решение'

    foreach($GroupDubleIncident in $GroupDubleIncidents)
    {
        $NumberOwnerIncident = $GroupDubleIncident.Group[0].'Решение'.substring(24).split(' ')[0]

        for($i=0; $i -lt $Incidents.Count; $i++)
        {
            if($Incidents[$i].Номер -like $NumberOwnerIncident)
            {
                $Points = $GroupDubleIncident.Group | where{$_.'Станция' -like $Incidents[$i].'Станция'}
                #$Incidents[$i]
                foreach($Point in $Points)
                {
                    if($Incidents[$i].'Узел' -cnotmatch $Point.'Узел')
                    {
                        $Incidents[$i].'Узел' += ", $($Point.'Узел')"
                    }
                }
                #$Incidents[$i]
            }
        }
    }

    return $Incidents
}

function sort-IncidentsInCategory()
{
    param($Incidents,
            $NumbersIncidensFromPreviousSmena)
    
    function get-SmenaTime ()
    {
        $Object=New-Object psobject
        $Object|Add-Member -MemberType NoteProperty -Name 'NowSmenaTime' -Value "" -Force
        $Object|Add-Member -MemberType NoteProperty -Name 'AllSmenaTime' -Value "" -Force

        $DateNow = (get-date).AddDays(0)

        if($DateNow -le (Get-Date -Hour 4 -Minute 0 -Second 0)){
            $Object.NowSmenaTime = (Get-Date -Hour 0 -Minute 0 -Second 0) - [timespan]("04:00")
            $Object.AllSmenaTime = (Get-Date -Hour 0 -Minute 0 -Second 0) - [timespan]("16:00")
        }
        if($DateNow -gt (Get-Date -Hour 4 -Minute 0 -Second 0) -and (get-date) -le (Get-Date -Hour 16 -Minute 0 -Second 0)){
            $Object.NowSmenaTime=Get-Date -Hour 8 -Minute 0 -Second 0
            $Object.AllSmenaTime=Get-Date -Hour 8 -Minute 0 -Second 0
        }
        if($DateNow -gt (Get-Date -Hour 16 -Minute 0 -Second 0)){
            $Object.NowSmenaTime=Get-Date -Hour 20 -Minute 0 -Second 0
            $Object.AllSmenaTime=Get-Date -Hour 8 -Minute 0 -Second 0
        }

        return $Object
    }

    function get-ObjectSmenaCategory()
    {
        $Object=New-Object psobject
        $Object|Add-Member -MemberType NoteProperty -Name 'Smena24Hours' -Value "" -Force
        $Object|Add-Member -MemberType NoteProperty -Name 'LastSmena' -Value "" -Force
        $Object|Add-Member -MemberType NoteProperty -Name 'ClosedIncidentsForSmena' -Value "" -Force
        $Object|Add-Member -MemberType NoteProperty -Name 'ToDoForNextSmena' -Value "" -Force
        $Object|Add-Member -MemberType NoteProperty -Name 'ToDoForNextSmenaUVST' -Value "" -Force
        $Object|Add-Member -MemberType NoteProperty -Name 'IncidensFromPreviousSmena' -Value "" -Force
        $Object|Add-Member -MemberType NoteProperty -Name 'ClosedIncidentsForSmena24Hours' -Value "" -Force
        $Object|Add-Member -MemberType NoteProperty -Name 'TI_Incidents24Hours' -Value "" -Force
        $Object|Add-Member -MemberType NoteProperty -Name 'TI_Incidents24Hours_WithProblem' -Value "" -Force
        $Object|Add-Member -MemberType NoteProperty -Name 'IncidensFromPreviousSmenaAuto' -Value "" -Force
        return $Object
    }

    $SmenaTime = get-SmenaTime
    $Report = get-ObjectSmenaCategory
    $Report.IncidensFromPreviousSmenaAuto = $Incidents|where{$_."Фактическое завершение" -gt $SmenaTime.NowSmenaTime `
            -and $_.'Время направления в работу' -lt $SmenaTime.NowSmenaTime -and $_."Рабочая группа" -like "*КРАСН"}
    $Report.Smena24Hours = $Incidents|where{$_."Время направления в работу" -gt $SmenaTime.AllSmenaTime}
    $Report.LastSmena = $Incidents|where{$_."Время направления в работу" -gt $SmenaTime.NowSmenaTime} 
    $Report.ClosedIncidentsForSmena = $Report.LastSmena|where{$_.Инфраструктурный -eq $true `
            -and $_."Рабочая группа" -like "*КРАСН" -and ($_."Статус" -match "Выполнен" -or $_."Статус" -match "Закрыт")}
    $Report.ClosedIncidentsForSmena24Hours = $Report.Smena24Hours|where{$_.Инфраструктурный -eq $true `
            -and $_."Рабочая группа" -like "*КРАСН" -and ($_."Статус" -match "Выполнен" -or $_."Статус" -match "Закрыт")}
    $Report.ToDoForNextSmena = $Incidents|where{$_."Рабочая группа" -like "*КРАСН" `
            -and ($_."Статус" -match "В работе" -or $_."Статус" -match "Приостановлен")}
    $Report.ToDoForNextSmenaUVST = $Incidents|where{$_."Рабочая группа" -like "*ЮВСТ" `
            -and ($_."Статус" -match "В работе" -or $_."Статус" -match "Приостановлен")}
    $Report.TI_Incidents24Hours = $Report.Smena24Hours | where{$_.'Технический' -eq $true}
    $Report.TI_Incidents24Hours_WithProblem = $Report.TI_Incidents24Hours|where{$_."Решение" -cmatch "ПРБ" `
            -and $_."Рабочая группа" -like "*КРАСН"}     ##"..[00-99]-[00000000-99999999]"}

    if($NumbersIncidensFromPreviousSmena)
    {
        $ObjPrevSmen=@()
        foreach($IncidentPreviousSmena in $NumbersIncidensFromPreviousSmena.split(","))
        {
            $ObjPrevSmen += $Incidents|where{$_."Рабочая группа" -like "*КРАСН" -and $IncidentPreviousSmena -contains $_."Номер"}
        }
        $Report.IncidensFromPreviousSmena = $ObjPrevSmen
    }

    return $Report
}

function get-IncedentsCounters()
{
    param ($LastSmenaReport)

    function get-ObjectForIncidentsCounters()
    {
        $Object=New-Object psobject
        $Object|Add-Member -MemberType NoteProperty -Name 'KSK_InfraInc' -Value 0 -Force
        $Object|Add-Member -MemberType NoteProperty -Name 'KSK_InfraIncToWork' -Value 0 -Force

        $Object|Add-Member -MemberType NoteProperty -Name 'KRS_TechnoInc' -Value 0 -Force
        $Object|Add-Member -MemberType NoteProperty -Name 'KRS_TechnoIncToWork' -Value 0 -Force

        $Object|Add-Member -MemberType NoteProperty -Name 'UVST_InfraInc' -Value 0 -Force
        $Object|Add-Member -MemberType NoteProperty -Name 'UVST_InfraIncToWork' -Value 0 -Force

        $Object|Add-Member -MemberType NoteProperty -Name 'UVST_TechnoInc' -Value 0 -Force
        $Object|Add-Member -MemberType NoteProperty -Name 'UVST_TechnoIncToWork' -Value 0 -Force

        $Object|Add-Member -MemberType NoteProperty -Name 'ZSIB_InfraInc' -Value 0 -Force
        $Object|Add-Member -MemberType NoteProperty -Name 'ZSIB_InfraIncToWork' -Value 0 -Force

        $Object|Add-Member -MemberType NoteProperty -Name 'ZSIB_TechnoInc' -Value 0 -Force
        $Object|Add-Member -MemberType NoteProperty -Name 'ZSIB_TechnoIncToWork' -Value 0 -Force

        return $Object
    }

    $IncedentsCounters = get-ObjectForIncidentsCounters

        for($i=0; $i -lt ($LastSmenaReport|measure).Count;$i++)
        {
            if($LastSmenaReport[$i]."Рабочая группа" -like "*КРАСН")
            {
                if($LastSmenaReport[$i].Технический  -eq $true)
                {
                    if($LastSmenaReport[$i]."Статус" -match "В работе" -or $LastSmenaReport[$i]."Статус" -match "Приостановлен")
                    {
                        $IncedentsCounters.KRS_TechnoIncToWork++
                    }
                        $IncedentsCounters.KRS_TechnoInc++
                }
                if($LastSmenaReport[$i].Инфраструктурный  -eq $true)
                {
                    if($LastSmenaReport[$i]."Статус" -match "В работе" -or $LastSmenaReport[$i]."Статус" -match "Приостановлен")
                    {
                        $IncedentsCounters.KSK_InfraIncToWork++
                    }
                        $IncedentsCounters.KSK_InfraInc++
                }
            
            }

            if($LastSmenaReport[$i]."Рабочая группа" -like "*ЮВСТ")
            {
                if($LastSmenaReport[$i].Технический  -eq $true)
                {
                    if($LastSmenaReport[$i]."Статус" -match "В работе" -or $LastSmenaReport[$i]."Статус" -match "Приостановлен")
                    {
                        $IncedentsCounters.UVST_TechnoIncToWork++
                    }
                        $IncedentsCounters.UVST_TechnoInc++
                }
                if($LastSmenaReport[$i].Инфраструктурный  -eq $true)
                {
                    if($LastSmenaReport[$i]."Статус" -match "В работе" -or $LastSmenaReport[$i]."Статус" -match "Приостановлен")
                    {
                        $IncedentsCounters.UVST_InfraIncToWork++
                    }
                        $IncedentsCounters.UVST_InfraInc++
                }
            
            }

            if($LastSmenaReport[$i]."Рабочая группа" -like "*ЗСИБ")
            {
                if($LastSmenaReport[$i].Технический  -eq $true)
                {
                    if($LastSmenaReport[$i]."Статус" -match "В работе" -or $LastSmenaReport[$i]."Статус" -match "Приостановлен")
                    {
                        $IncedentsCounters.ZSIB_TechnoIncToWork++
                    }
                        $IncedentsCounters.ZSIB_TechnoInc++
                }
                if($LastSmenaReport[$i].Инфраструктурный  -eq $true)
                {
                    if($LastSmenaReport[$i]."Статус" -match "В работе" -or $LastSmenaReport[$i]."Статус" -match "Приостановлен")
                    {
                        $IncedentsCounters.ZSIB_InfraIncToWork++
                    }
                        $IncedentsCounters.ZSIB_InfraInc++
                }
            
            }
        }

        return $IncedentsCounters
        <#
        echo "КРАСН: ТИ - $KRS_TechnoInc ИИ - $KSK_InfraInc"
        echo ""
        echo "ЮВСТ: ТИ - $UVST_TechnoInc ИИ - $UVST_InfraInc"
        echo ""
        echo "ЗСИБ: ТИ - $ZSIB_TechnoInc ИИ - $ZSIB_InfraInc"
    #>
}

function get-FinalReport()
{
    param ($Report,
            $IncedentsCountersForLastSmena,
            $IncedentsCountersForSmena24Hours,
            $IncedentsCountersForAll,
            $value)

    function get-FormatReport($Report){
        $FormatReport=@()
        for($i=0;$i -lt ($Report|measure).Count;$i++){
            $StartTime = ""
            $EndTime = ""
            $Comment = ""
            $Tech = ""
            if($Report[$i].Технический -eq $true){$Tech = "(ТИ)"}
            if($Report[$i]."Время направления в работу" -notlike ""){$StartTime = (get-date($Report[$i]."Время направления в работу" - [timespan]("04:00")) -Format "HH:mm")}
            if($Report[$i]."Время закрытия события в системе мониторинга"-notlike ""){$EndTime = (get-date($Report[$i]."Время закрытия события в системе мониторинга" - [timespan]("04:00")) -Format "HH:mm")}
            if($Report[$i]."Решение" -match "..[00-99]-[00000000-99999999]" -or $Report[$i]."Решение" -match "возможно" -or $Report[$i]."Корневая причина" -like ""){$Comment = $Report[$i]."Решение"} #-or $Report[$i]."Решение" -match "ПРБ[00-99]-[00000000-99999999]"
            #filter for coment
            if($Comment -cmatch "ДО:" -and $Comment -cmatch " ПОСЛЕ:"){$Comment = $Comment.Substring(0,($Comment.IndexOf('ДО:')))}
            #190721#$FormatReport += $StartTime + " - " + $EndTime  + " " + $Tech + $Report[$i]."Номер" + " " + $Report[$i]."Подробное описание" + ". " + $Report[$i]."Корневая причина" + ". " + $Comment
            $FormatReport += $StartTime + " - " + $EndTime  + " " + $Tech + $Report[$i]."Номер" + `
                ". Ст. " + $Report[$i]."Станция" + ", " + $Report[$i]."Узел" + ". " + `
                $Report[$i]."Корневая причина" + ". " + $Comment
        }
        return $FormatReport
    }

    function Get-CountersInfAndTech()
    {
        param($IncedentsMass)
        
        $CountInfrastructure = (($IncedentsMass | where{$_.Инфраструктурный  -eq $true}) | measure).Count
        $CountTechnical = (($IncedentsMass |  where{$_.Инфраструктурный  -eq $true}) | measure).Count

        return $CountInfrastructure, $CountTechnical
    }

    function get-FullReport()
    {
        $CountInfrastructureFromPreviousSmenaAuto, $CountTechnicalFromPreviousSmenaAuto = `
                Get-CountersInfAndTech ($Report.IncidensFromPreviousSmenaAuto | where{$_."Статус" -match "Выполнен" -or $_."Статус" -match "Закрыт"})

        $TMPText=@()
        $TMPText += "1. Инциденты, переданные по смене (решенные).`n"
        $TMPText += get-FormatReport ($Report.IncidensFromPreviousSmenaAuto | where{$_."Статус" -match "Выполнен" -or $_."Статус" -match "Закрыт"})
        $TMPText += "`n2. Инциденты, переданные по смене (не решенные).`n"
        $TMPText += get-FormatReport ($Report.IncidensFromPreviousSmenaAuto | where{$_."Статус" -match "В работе" -or $_."Статус" -match "Приостановлен"})
        $TMPText += "`n3. Инциденты по питанию.`n"
        $TMPText += get-FormatReport ($Report.ClosedIncidentsForSmena | where{$_.Электроэнергия -eq $true})
        $TMPText += "`n4. Инциденты по СПД.`n"
        $TMPText += get-FormatReport ($Report.ClosedIncidentsForSmena | where{$_.Электроэнергия -eq $false})
        $TMPText += "`n5. Инциденты в работе."
        $TMPText += get-FormatReport $Report.ToDoForNextSmena
        $TMPText += "`n6. Инциденты в работе ЮВСТ."
        $TMPText += get-FormatReport $Report.ToDoForNextSmenaUVST
        $TMPText += "`n`nКол-во решенных переданных ИНЦ КРАСН:."
        $TMPText +=  "ТИ - "+$CountTechnicalFromPreviousSmenaAuto + " ИИ - " + $CountInfrastructureFromPreviousSmenaAuto
        $TMPText += "`n`nЗа смену обработано:"
        $TMPText += "КРАСН: ТИ - "+$IncedentsCountersForLastSmena.KRS_TechnoInc + " ИИ - " + $IncedentsCountersForLastSmena.KSK_InfraInc
        #$TMPText += "В работе: ТИ - "+$IncedentsCountersForLastSmena.KRS_TechnoIncToWork + " ИИ - " + $IncedentsCountersForLastSmena.KSK_InfraIncToWork
        $TMPText += "В работе: ТИ - "+$IncedentsCountersForAll.KRS_TechnoIncToWork + " ИИ - " + $IncedentsCountersForAll.KSK_InfraIncToWork
        $TMPText += "`nЮВСТ: ТИ - "+$IncedentsCountersForLastSmena.UVST_TechnoInc + " ИИ - " + $IncedentsCountersForLastSmena.UVST_InfraInc
        #$TMPText += "В работе: ТИ - "+$IncedentsCountersForLastSmena.UVST_TechnoIncToWork + " ИИ - " + $IncedentsCountersForLastSmena.UVST_InfraIncToWork
        $TMPText += "В работе: ТИ - "+$IncedentsCountersForAll.UVST_TechnoIncToWork + " ИИ - " + $IncedentsCountersForAll.UVST_InfraIncToWork
        $TMPText += "`nЗСИБ: ТИ - "+$IncedentsCountersForLastSmena.ZSIB_TechnoInc + " ИИ - " + $IncedentsCountersForLastSmena.ZSIB_InfraInc
        #$TMPText += "В работе: ТИ - "+$IncedentsCountersForLastSmena.ZSIB_TechnoIncToWork + " ИИ - " + $IncedentsCountersForLastSmena.ZSIB_InfraIncToWork
        $TMPText += "В работе: ТИ - "+$IncedentsCountersForAll.ZSIB_TechnoIncToWork + " ИИ - " + $IncedentsCountersForAll.ZSIB_InfraIncToWork

        return $TMPText
    }

    function get-FullEndSmenaReport()
    {
        $Energo = 'Аварийное отключение электроэнергии','Плановые работы энергетиков, без оповещения ИВЦ'
        $TMPText=@()
        $TMPText += "1. Инциденты, переданные по смене (решенные).`n"
        $TMPText += get-FormatReport ($Report.IncidensFromPreviousSmena | where{$_."Статус" -match "Выполнен" -or $_."Статус" -match "Закрыт"})
        $TMPText += "`n2. Инциденты, переданные по смене (не решенные).`n"
        $TMPText += get-FormatReport ($Report.IncidensFromPreviousSmena | where{$_."Статус" -match "В работе" -or $_."Статус" -match "Приостановлен"})
        #$Energo='Аварийное отключение','Плановые работы энергетиков'
        <#$Energo = 'Аварийное отключение электроэнергии','Плановые работы энергетиков, без оповещения ИВЦ'
        $TMPText += get-FormatReport ($Report.ClosedIncidentsForSmena24Hours | where{$Energo -contains $_."Корневая причина"})
        $TMPText += "3.	Инциденты СПД.`n"
        $TMPText += get-FormatReport ($Report.ClosedIncidentsForSmena24Hours | where{$Energo -notcontains $_."Корневая причина"})
        #>
        $TMPText += "`n3. Инциденты по питанию.`n"
        $TMPText += get-FormatReport ($Report.ClosedIncidentsForSmena24Hours | where{$_.Электроэнергия -eq $true})
        $TMPText += "`n4. Инциденты по СПД.`n"
        $TMPText += get-FormatReport ($Report.ClosedIncidentsForSmena24Hours | where{$_.Электроэнергия -eq $false})
        $TMPText += "`n5. Инциденты в работе."
        $TMPText += get-FormatReport $Report.ToDoForNextSmena
        $TMPText += "`n6. Инциденты в работе ЮВСТ."
        $TMPText += get-FormatReport $Report.ToDoForNextSmenaUVST
        $TMPText += "`n`nЗа сутки обработано:"
        $TMPText += "КРАСН: ТИ - "+$IncedentsCountersForSmena24Hours.KRS_TechnoInc + " ИИ - " + $IncedentsCountersForSmena24Hours.KSK_InfraInc
        $TMPText += "`nЮВСТ: ТИ - "+$IncedentsCountersForSmena24Hours.UVST_TechnoInc + " ИИ - " + $IncedentsCountersForSmena24Hours.UVST_InfraInc
        $TMPText += "`nЗСИБ: ТИ - "+$IncedentsCountersForSmena24Hours.ZSIB_TechnoInc + " ИИ - " + $IncedentsCountersForSmena24Hours.ZSIB_InfraInc

        $TMPText += "`n`nИнтересные ТИ за сутки:"
        $TMPText += get-FormatReport $Report.TI_Incidents24Hours_WithProblem
        

        return $TMPText
    }
    $Result1 = get-FullReport
    $Result2 = get-FullEndSmenaReport

    return $Result1,$Result2
}

function New-Incidets()
{
    param ($Incidents,$NumbersIncidentsSoonComplite)

    #($Incidents|where{$_.'Статус' -like '4-Выполнен'}).count
    for($i=0; $i -lt $Incidents.Count; $i++)
    {
        if($NumbersIncidentsSoonComplite.split(',') -contains $Incidents[$i].'Номер')
        {
            #$Incidents[$i].'Номер'
            $Incidents[$i].'Статус' = '4-Выполнен'
        }
    }
    #($Incidents|where{$_.'Статус' -like '4-Выполнен'}).count
    return $Incidents
}

function get-SmenaReport()
{
    <#
    .Synopsis
        Подготовка отчета смены
    .Description
        Подготовка отчета смены из экспортируемого файла ЕСПП
    .Parameter InputFile
        Месторасположение экспортируемого файла из ЕПП (по умолчанию ищет в папке Downloads)
    .Parameter NumbersIncidensFromPreviousSmena
        Перечисленные инциденты, переданные по смене
    .Parameter NumbersIncidentsSoonComplite
        Перечисленные инциденты, которые будут закрыты при передачи смены
    .Example
        get-SmenaReport -InputFile 'D:\Reports\Export.txt'
        
        get-SmenaReport -NumbersIncidensFromPreviousSmena 'ИНЦ19-00331826,ИНЦ19-00331877'
    
    .Inputs
        System.String
    #>

    param ($InputFile,$NumbersIncidensFromPreviousSmena,$NumbersIncidentsSoonComplite,$AddDublicatePoint)
    #if($IncidetsToWork -and $IncidetsToWork -match ",ИНЦ"){$IncidetsToWork=$IncidetsToWork.slpit(',')}

    if(!(Test-Path "$env:USERPROFILE\Documents\SPD_Reports")){New-Item -ItemType Directory "$env:USERPROFILE\Documents\SPD_Reports"}

    $Incidents = get-ReportFile -InputFile $InputFile | Sort-Object "Электроэнергия", "Время направления в работу"  
    if($NumbersIncidentsSoonComplite)
    {
        $Incidents = New-Incidets -Incidents $Incidents -NumbersIncidentsSoonComplite $NumbersIncidentsSoonComplite
        $NumbersIncidentsSoonComplite > "$env:USERPROFILE\Documents\SPD_Reports\NumbersIncidentsSoonComplite.txt"
    }
    if($AddDublicatePoint)
    {
        $Incidents = add-StationsWhereDublicateInc $Incidents
    }
    $Report = sort-IncidentsInCategory -Incidents $Incidents -NumbersIncidensFromPreviousSmena $NumbersIncidensFromPreviousSmena
    $IncedentsCountersForLastSmena = get-IncedentsCounters $Report.LastSmena
    $IncedentsCountersForSmena24Hours = get-IncedentsCounters $Report.Smena24Hours
    $IncedentsCountersForAll = get-IncedentsCounters $Incidents
    $FinalReport,$FinalReportFull = Get-FinalReport $Report $IncedentsCountersForLastSmena $IncedentsCountersForSmena24Hours $IncedentsCountersForAll
    $date = get-date -Format "MM-dd-HH-mm"
    
    $FinalReport  > "$env:USERPROFILE\Documents\SPD_Reports\$date-report.txt"
    $FinalReportFull  > "$env:USERPROFILE\Documents\SPD_Reports\$date-Fullreport.txt"
    $Incidents|Sort-Object "Время направления в работу" -Descending|Export-Clixml "$env:USERPROFILE\Documents\SPD_Reports\Base.xml"
    $Incidents|Sort-Object "Время направления в работу" -Descending|Export-Csv "$env:USERPROFILE\Documents\SPD_Reports\Base.csv"
    Get-ChildItem "$env:USERPROFILE\Documents\SPD_Reports\" | Sort-Object LastWriteTime -Descending | Select-Object -Skip 15 |Remove-Item
    write-errorlog
    #$FinalReport
}

function write-errorlog()
{
    $ErrDir = "$env:USERPROFILE\Documents\SPD_Reports\ErrDir"
    $ErrFile = "GetSmenaReport_ErrorLog.log"

    if(!(Test-Path $ErrDir)){New-Item -ItemType Directory $ErrDir}
    if($Error)
    {
        echo "_______________" >> "$ErrDir\$ErrFile"
        Get-Date >> "$ErrDir\$ErrFile"
        $Error >> "$ErrDir\$ErrFile"
    }
}

function start-smenaform()
{
    powershell -WindowStyle Hidden -file "C:\windows\System32\WindowsPowerShell\v1.0\Modules\GetSmenaReport\StartSmenaRep.ps1" 
}


#start-smenaform
get-SmenaReport
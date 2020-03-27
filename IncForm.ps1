param($IncNum)
$Error.Clear()

function Start-IncForm()
{
    param($IncData)

    Add-Type -AssemblyName System.Windows.Forms

    
    $label_Incident = new-object System.Windows.Forms.Label
    $label_IncNumber = new-object System.Windows.Forms.Label
    $Label_WorkGroup = new-object System.Windows.Forms.Label
    $label_Worker = new-object System.Windows.Forms.Label
    $label_Station = new-object System.Windows.Forms.Label
    $textBox_IncNumber = new-object System.Windows.Forms.TextBox
    $textBox_WorkGroup = new-object System.Windows.Forms.TextBox
    $radioButtonInfra = new-object System.Windows.Forms.RadioButton
    $radioButtonTech = new-object System.Windows.Forms.RadioButton
    $radioButtonHuman = new-object System.Windows.Forms.RadioButton
    $textBox_Worker = new-object System.Windows.Forms.TextBox
    $textBox_Station = new-object System.Windows.Forms.TextBox
    $label_Point = new-object System.Windows.Forms.Label
    $textBox_Point = new-object System.Windows.Forms.TextBox
    $label_TimeStart = new-object System.Windows.Forms.Label
    $label_TimeMonitor = new-object System.Windows.Forms.Label
    $label_TimeEnd = new-object System.Windows.Forms.Label
    $label_Status = new-object System.Windows.Forms.Label
    $textBox_Status = new-object System.Windows.Forms.TextBox
    $textBox_TimeStart = new-object System.Windows.Forms.TextBox
    $textBox_TimeMonitor = new-object System.Windows.Forms.TextBox
    $textBox_TimeEnd = new-object System.Windows.Forms.TextBox
    $label_IP = new-object System.Windows.Forms.Label
    $textBox_IP = new-object System.Windows.Forms.TextBox
    $label_DN = new-object System.Windows.Forms.Label
    $textBox_DN = new-object System.Windows.Forms.TextBox
    $label_LongInformation = new-object System.Windows.Forms.Label
    $textBox_LongInformation = new-object System.Windows.Forms.TextBox
    $label_Solved = new-object System.Windows.Forms.Label
    $textBox_Solved = new-object System.Windows.Forms.TextBox
    $SuspendLayout
    ## 
    ## label_Incident
    ## 
    $label_Incident.AutoSize = $true
    $label_Incident.Location = new-object System.Drawing.Point(235, 9)
    $label_Incident.Name = "label_Incident"
    $label_Incident.Size = new-object System.Drawing.Size(56, 13)
    $label_Incident.TabIndex = 0
    $label_Incident.Text = "Инцидент"
    ## 
    ## label_IncNumber
    ## 
    $label_IncNumber.AutoSize = $true
    $label_IncNumber.Location = new-object System.Drawing.Point(22, 61)
    $label_IncNumber.Name = "label_IncNumber"
    $label_IncNumber.Size = new-object System.Drawing.Size(41, 13)
    $label_IncNumber.TabIndex = 0
    $label_IncNumber.Text = "Номер"
    ## 
    ## Label_WorkGroup
    ## 
    $Label_WorkGroup.AutoSize = $true
    $Label_WorkGroup.Location = new-object System.Drawing.Point(22, 113)
    $Label_WorkGroup.Name = "Label_WorkGroup"
    $Label_WorkGroup.Size = new-object System.Drawing.Size(86, 13)
    $Label_WorkGroup.TabIndex = 0
    $Label_WorkGroup.Text = "Рабочая группа"
    ## 
    ## label_Worker
    ## 
    $label_Worker.AutoSize = $true
    $label_Worker.Location = new-object System.Drawing.Point(22, 137)
    $label_Worker.Name = "label_Worker"
    $label_Worker.Size = new-object System.Drawing.Size(74, 13)
    $label_Worker.TabIndex = 0
    $label_Worker.Text = "Исполнитель"
    ## 
    ## label_Station
    ## 
    $label_Station.AutoSize = $true
    $label_Station.Location = new-object System.Drawing.Point(22, 161)
    $label_Station.Name = "label_Station"
    $label_Station.Size = new-object System.Drawing.Size(49, 13)
    $label_Station.TabIndex = 0
    $label_Station.Text = "Станция"
    ## 
    ## radioButtonInfra
    ## 
    $radioButtonInfra.AutoSize = $true
    $radioButtonInfra.Location = new-object System.Drawing.Point(25, 31)
    $radioButtonInfra.Name = "radioButtonInfra"
    $radioButtonInfra.Size = new-object System.Drawing.Size(123, 17)
    $radioButtonInfra.TabIndex = 2
    $radioButtonInfra.TabStop = $true
    $radioButtonInfra.Text = "Инфраструктурный"
    $radioButtonInfra.UseVisualStyleBackColor = $true
    if($IncData.Инфраструктурный -eq $true){$radioButtonInfra.Checked=$true}    
    ## 
    ## radioButtonTech
    ## 
    $radioButtonTech.AutoSize = $true
    $radioButtonTech.Location = new-object System.Drawing.Point(219, 31)
    $radioButtonTech.Name = "radioButtonTech"
    $radioButtonTech.Size = new-object System.Drawing.Size(90, 17)
    $radioButtonTech.TabIndex = 2
    $radioButtonTech.TabStop = $true
    $radioButtonTech.Text = "Технический"
    $radioButtonTech.UseVisualStyleBackColor = $true
    if($IncData.Технический -eq $true){$radioButtonTech.Checked=$true}    
    ## 
    ## radioButtonHuman
    ## 
    $radioButtonHuman.AutoSize = $true
    $radioButtonHuman.Location = new-object System.Drawing.Point(400, 31)
    $radioButtonHuman.Name = "radioButtonHuman"
    $radioButtonHuman.Size = new-object System.Drawing.Size(91, 17)
    $radioButtonHuman.TabIndex = 2
    $radioButtonHuman.TabStop = $true
    $radioButtonHuman.Text = "Человечачий"
    $radioButtonHuman.UseVisualStyleBackColor = $true
    if($IncData.Технический -eq $false -and $IncData.Инфраструктурный -eq $false){$radioButtonHuman.Checked=$true}  

    ## 
    ## label_Point
    ## 
    $label_Point.AutoSize = $true
    $label_Point.Location = new-object System.Drawing.Point(22, 187)
    $label_Point.Name = "label_Point"
    $label_Point.Size = new-object System.Drawing.Size(33, 13)
    $label_Point.TabIndex = 0
    $label_Point.Text = "Узел"
    ## 
    ## label_TimeStart
    ## 
    $label_TimeStart.AutoSize = $true
    $label_TimeStart.Location = new-object System.Drawing.Point(22, 254)
    $label_TimeStart.Name = "label_TimeStart"
    $label_TimeStart.Size = new-object System.Drawing.Size(148, 13)
    $label_TimeStart.TabIndex = 0
    $label_TimeStart.Text = "Время направленя в группу"
    ## 
    ## label_TimeMonitor
    ## 
    $label_TimeMonitor.AutoSize = $true
    $label_TimeMonitor.Location = new-object System.Drawing.Point(195, 255)
    $label_TimeMonitor.Name = "label_TimeMonitor"
    $label_TimeMonitor.Size = new-object System.Drawing.Size(170, 13)
    $label_TimeMonitor.TabIndex = 0
    $label_TimeMonitor.Text = "Время закрытия в мониторинге"
    ## 
    ## label_TimeEnd
    ## 
    $label_TimeEnd.AutoSize = $true
    $label_TimeEnd.Location = new-object System.Drawing.Point(379, 255)
    $label_TimeEnd.Name = "label_TimeEnd"
    $label_TimeEnd.Size = new-object System.Drawing.Size(141, 13)
    $label_TimeEnd.TabIndex = 0
    $label_TimeEnd.Text = "Фактическое завершение"
    ## 
    ## label_Status
    ## 
    $label_Status.AutoSize = $true
    $label_Status.Location = new-object System.Drawing.Point(22, 87)
    $label_Status.Name = "label_Status"
    $label_Status.Size = new-object System.Drawing.Size(41, 13)
    $label_Status.TabIndex = 0
    $label_Status.Text = "Статус"
    ## 
    ## label_IP
    ## 
    $label_IP.AutoSize = $true
    $label_IP.Location = new-object System.Drawing.Point(22, 211)
    $label_IP.Name = "label_IP"
    $label_IP.Size = new-object System.Drawing.Size(17, 13)
    $label_IP.TabIndex = 0
    $label_IP.Text = "IP"


    ##############################################################


    ## 
    ## textBox_Worker
    ## 
    $textBox_Worker.Location = new-object System.Drawing.Point(198, 134)
    $textBox_Worker.Name = "textBox_Worker"
    $textBox_Worker.Size = new-object System.Drawing.Size(341, 20)
    $textBox_Worker.TabIndex = 1
    $textBox_Worker.Text = $IncData.Исполнитель
    ## 
    ## textBox_Station
    ## 
    $textBox_Station.Location = new-object System.Drawing.Point(198, 158)
    $textBox_Station.Name = "textBox_Station"
    $textBox_Station.Size = new-object System.Drawing.Size(341, 20)
    $textBox_Station.TabIndex = 1
    $textBox_Station.Text = $IncData.Станция
    ## 
    ## textBox_Point
    ## 
    $textBox_Point.Location = new-object System.Drawing.Point(198, 184)
    $textBox_Point.Name = "textBox_Point"
    $textBox_Point.Size = new-object System.Drawing.Size(341, 20)
    $textBox_Point.TabIndex = 1
    $textBox_Point.Text = $IncData.Узел
    ## 
    ## textBox_IncNumber
    ## 
    $textBox_IncNumber.Location = new-object System.Drawing.Point(198, 58)
    $textBox_IncNumber.Name = "textBox_IncNumber"
    $textBox_IncNumber.Size = new-object System.Drawing.Size(341, 20)
    $textBox_IncNumber.TabIndex = 1
    $textBox_IncNumber.Text=$IncData.'Номер'
    ## 
    ## textBox_WorkGroup
    ## 
    $textBox_WorkGroup.Location = new-object System.Drawing.Point(198, 110)
    $textBox_WorkGroup.Name = "textBox_WorkGroup"
    $textBox_WorkGroup.Size = new-object System.Drawing.Size(341, 20)
    $textBox_WorkGroup.TabIndex = 1
    $textBox_WorkGroup.Text = $IncData.'Рабочая группа'
    ## 
    ## textBox_Status
    ## 
    $textBox_Status.Location = new-object System.Drawing.Point(198, 84)
    $textBox_Status.Name = "textBox_Status"
    $textBox_Status.Size = new-object System.Drawing.Size(341, 20)
    $textBox_Status.TabIndex = 1
    $textBox_Status.Text = $IncData.Статус
    ## 
    ## textBox_TimeStart
    ## 
    $textBox_TimeStart.Location = new-object System.Drawing.Point(24, 270)
    $textBox_TimeStart.Name = "textBox_TimeStart"
    $textBox_TimeStart.Size = new-object System.Drawing.Size(156, 20)
    $textBox_TimeStart.TabIndex = 1
    $textBox_TimeStart.Text = $IncData.'Время направления в работу'
    ## 
    ## textBox_TimeMonitor
    ## 
    $textBox_TimeMonitor.Location = new-object System.Drawing.Point(198, 270)
    $textBox_TimeMonitor.Name = "textBox_TimeMonitor"
    $textBox_TimeMonitor.Size = new-object System.Drawing.Size(167, 20)
    $textBox_TimeMonitor.TabIndex = 1
    $textBox_TimeMonitor.Text = $IncData.'Время закрытия события в системе мониторинга'
    ## 
    ## textBox_TimeEnd
    ## 
    $textBox_TimeEnd.Location = new-object System.Drawing.Point(382, 270)
    $textBox_TimeEnd.Name = "textBox_TimeEnd"
    $textBox_TimeEnd.Size = new-object System.Drawing.Size(157, 20)
    $textBox_TimeEnd.TabIndex = 1
    $textBox_TimeEnd.Text = $IncData.'Фактическое завершение'
    ## 
    ## textBox_IP
    ## 
    $textBox_IP.Location = new-object System.Drawing.Point(198, 208)
    $textBox_IP.Name = "textBox_IP"
    $textBox_IP.Size = new-object System.Drawing.Size(341, 20)
    $textBox_IP.TabIndex = 1
    ## 
    ## label_DN
    ## 
    $label_DN.AutoSize = $true
    $label_DN.Location = new-object System.Drawing.Point(22, 235)
    $label_DN.Name = "label_DN"
    $label_DN.Size = new-object System.Drawing.Size(23, 13)
    $label_DN.TabIndex = 0
    $label_DN.Text = "DN"
    ## 
    ## textBox_DN
    ## 
    $textBox_DN.Location = new-object System.Drawing.Point(198, 232)
    $textBox_DN.Name = "textBox_DN"
    $textBox_DN.Size = new-object System.Drawing.Size(341, 20)
    $textBox_DN.TabIndex = 1
    ## 
    ## label_LongInformation
    ## 
    $label_LongInformation.AutoSize = $true
    $label_LongInformation.Location = new-object System.Drawing.Point(22, 302)
    $label_LongInformation.Name = "label_LongInformation"
    $label_LongInformation.Size = new-object System.Drawing.Size(114, 13)
    $label_LongInformation.TabIndex = 0
    $label_LongInformation.Text = "Подробное описание"
    ## 
    ## textBox_LongInformation
    ## 
    $textBox_LongInformation.Location = new-object System.Drawing.Point(24, 318)
    $textBox_LongInformation.Multiline = $true
    $textBox_LongInformation.Name = "textBox_LongInformation"
    $textBox_LongInformation.Size = new-object System.Drawing.Size(515, 218)
    $textBox_LongInformation.TabIndex = 1
    $textBox_LongInformation.Text = $IncData.'Подробное описание'
    ## 
    ## label_Solved
    ## 
    $label_Solved.AutoSize = $true
    $label_Solved.Location = new-object System.Drawing.Point(22, 539)
    $label_Solved.Name = "label_Solved"
    $label_Solved.Size = new-object System.Drawing.Size(52, 13)
    $label_Solved.TabIndex = 0
    $label_Solved.Text = "Решение"
    ## 
    ## textBox_Solved
    ## 
    $textBox_Solved.Location = new-object System.Drawing.Point(24, 555)
    $textBox_Solved.Multiline = $true
    $textBox_Solved.Name = "textBox_Solved"
    $textBox_Solved.Size = new-object System.Drawing.Size(515, 67)
    $textBox_Solved.TabIndex = 1
    $textBox_Solved.Text = $IncData.Решение
    ## 
    ## Form_Incident
    ## 
 
    [System.Windows.Forms.Form]$form2 = New-Object System.Windows.Forms.Form;
    $form2.Width = 111;
    $form2.Height = 319;
    $form2.ClientSize = new-object System.Drawing.Size(551, 632)
    $form2.Controls.Add($radioButtonHuman)
    $form2.Controls.Add($radioButtonTech)
    $form2.Controls.Add($radioButtonInfra)
    $form2.Controls.Add($textBox_TimeEnd)
    $form2.Controls.Add($textBox_TimeMonitor)
    $form2.Controls.Add($textBox_Solved)
    $form2.Controls.Add($textBox_LongInformation)
    $form2.Controls.Add($textBox_TimeStart)
    $form2.Controls.Add($textBox_DN)
    $form2.Controls.Add($textBox_IP)
    $form2.Controls.Add($textBox_Point)
    $form2.Controls.Add($textBox_Station)
    $form2.Controls.Add($textBox_Worker)
    $form2.Controls.Add($textBox_Status)
    $form2.Controls.Add($textBox_WorkGroup)
    $form2.Controls.Add($label_TimeEnd)
    $form2.Controls.Add($label_DN)
    $form2.Controls.Add($label_TimeMonitor)
    $form2.Controls.Add($label_IP)
    $form2.Controls.Add($label_Solved)
    $form2.Controls.Add($label_LongInformation)
    $form2.Controls.Add($label_TimeStart)
    $form2.Controls.Add($label_Point)
    $form2.Controls.Add($textBox_IncNumber)
    $form2.Controls.Add($label_Station)
    $form2.Controls.Add($label_Status)
    $form2.Controls.Add($label_Worker)
    $form2.Controls.Add($Label_WorkGroup)
    $form2.Controls.Add($label_IncNumber)
    $form2.Controls.Add($label_Incident)
    $form2.Name = "Form_Incident"
    $form2.Text = "Form_Incident"



    $form2.ShowDialog() | Out-Null;
    
    
    write-errorlog
}

function write-errorlog()
{
    $ErrDir = "$env:USERPROFILE\Documents\SPD_Reports\ErrDir"
    $ErrFile = "IncForm_ErrorLog.log"

    if(!(Test-Path $ErrDir)){New-Item -ItemType Directory $ErrDir}
    if($Error)
    {
        echo "_______________" >> "$ErrDir\$ErrFile"
        Get-Date >> "$ErrDir\$ErrFile"
        $Error >> "$ErrDir\$ErrFile"
    }
}

$ImportData=Import-Clixml "$env:USERPROFILE\Documents\SPD_Reports\Base.xml"
$IncData = $ImportData|where{$_."Номер" -match $IncNum}
$IncData
Start-IncForm $IncData
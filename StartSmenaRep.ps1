Import-Module GetSmenaReport
$Error.Clear()



function Fill-IncidentsDataGridView(){
    $ImportData = Import-Clixml "$env:USERPROFILE\Documents\SPD_Reports\Base.xml"
    $IncidentsDataGridView.DataSource = [System.Collections.ArrayList]$ImportData
    $IncidentsDataGridView.RowHeadersVisible=$false
    
    $comboBox3Station.Items.Clear()
    $comboBox3Station.Items.AddRange(($ImportData|Group-Object "Станция").Name)
    $comboBox3Point.Items.Clear()
    $comboBox3Point.Items.AddRange(($ImportData|Group-Object "Узел").Name)
    $comboBox3Zone.Items.Clear()
    $comboBox3Zone.Items.AddRange(($ImportData|Group-Object "Рабочая группа").Name)

    #Sorting in process....
    <#.........
    $IncidentsDataGridView_ColumnHeaderMouseClick=[System.Windows.Forms.DataGridViewCellMouseEventHandler]{
        #Event Argument: $_ = [System.Windows.Forms.DataGridViewCellMouseEventArgs]
        #if ($IncidentsDataGridView.DataSource -is [System.Data.DataTable])
        #{
            $column = $IncidentsDataGridView.Columns[$_.ColumnIndex]
            $direction = [System.ComponentModel.ListSortDirection]::Ascending
        
            if ($column.HeaderCell.SortGlyphDirection -eq 'Descending')
            {
                $direction = [System.ComponentModel.ListSortDirection]::Descending
            }
        
            $IncidentsDataGridView.Sort($IncidentsDataGridView.Columns[$_.ColumnIndex], $direction)
        #}
    }
    #>
}

function grid_Click()
{
    $rowIndex = $IncidentsDataGridView.CurrentRow.Index
    $columnIndex = $IncidentsDataGridView.CurrentCell.ColumnIndex
    
    powershell "C:\windows\System32\WindowsPowerShell\v1.0\Modules\GetSmenaReport\IncForm.ps1" -IncNum $IncidentsDataGridView.Rows[$rowIndex].Cells[0].value
    #powershell "D:\Scripts\SmenaMZFK\IncForm.ps1" -IncNum $IncidentsDataGridView.Rows[$rowIndex].Cells[0].value

    <#
    Write-Host $rowIndex
    Write-Host $columnIndex 
    Write-Host $IncidentsDataGridView.Rows[$rowIndex]
    Write-Host $IncidentsDataGridView.Rows[$rowIndex].Cells[0].value
    Write-Host $IncidentsDataGridView.Rows[$rowIndex].Cells[$columnIndex].value
    #>
}

function button4_ECHZONE_Click()
{
    if($textBox4_ECHZone.Text -notlike "")
    {
        if(!(Test-Path "$env:USERPROFILE\Documents\SPD_Reports\Data")){New-Item -ItemType Directory -Path "$env:USERPROFILE\Documents\SPD_Reports\Data"}  #"$env:USERPROFILE\Documents\SPD_Reports\Data" }
        $textBox4_ECHZone.text > "$env:USERPROFILE\Documents\SPD_Reports\Data\ECH_Notificators.txt"

        powershell "C:\windows\System32\WindowsPowerShell\v1.0\Modules\GetSmenaReport\Form_ElectronicNotofocators.ps1"
        
    }
}

function find-incidents()
{
    param($SubstringFindObject,
            $PropertyFindObject)

    $IsDateTime = $false
    

    switch($PropertyFindObject)
    {
        "textBox3IncNumber" {$PropertyFindObject="Номер"}
        "comboBox3Zone" {$PropertyFindObject="Рабочая группа"}
        "comboBox3Station" {$PropertyFindObject="Станция"}
        "comboBox3Point" {$PropertyFindObject="Узел"}
        "textBox3Solved" {$PropertyFindObject="Решение"}
        "textBox3IP" {$PropertyFindObject="Краткое описание"}
        "dateTimePicker3TimeStart" {$PropertyFindObject="Время направления в работу"; $IsDateTime = $True; $IsTimeStart = $True}
        "dateTimePicker3TimeEnd" {$PropertyFindObject="Время направления в работу"; $IsDateTime = $True; $IsTimeStart = $false}
    }



    $IncidentsDataGridView.CurrentCell=$null
    $VisibleGrid=@()
    if($SubstringFindObject.length -gt 4)
    { 
        for($i=0;$i -lt $IncidentsDataGridView.ColumnCount;$i++)
        {
            if($IncidentsDataGridView.Columns[$i].Name -like $PropertyFindObject)
            {
                #Ecли строка не дата
                if($IsDateTime -eq $false)
                {
                    for($j=0;$j -lt $IncidentsDataGridView.RowCount; $j++)
                    {

                        if($IncidentsDataGridView.Rows[$j].Cells[$i].value -notmatch $SubstringFindObject)
                        {
                            $IncidentsDataGridView.Rows[$j].visible=$false
                        }else{
                            $IncidentsDataGridView.Rows[$j].visible=$True
                            $VisibleGrid+=$IncidentsDataGridView.Rows[$j].DataBoundItem
                        }
                    }
                }
                #Ecли строка является датой
                if($IsDateTime)
                {
                    for($j=0;$j -lt $IncidentsDataGridView.RowCount; $j++)
                    {
                            if($IsTimeStart)
                            {
                                if($IncidentsDataGridView.Rows[$j].Cells[$i].value -lt $SubstringFindObject)
                                {
                                    $IncidentsDataGridView.Rows[$j].visible=$false
                                }else{
                                    $IncidentsDataGridView.Rows[$j].visible=$True
                                    $VisibleGrid+=$IncidentsDataGridView.Rows[$j].DataBoundItem
                                }
                            }else{
                                if($IncidentsDataGridView.Rows[$j].Cells[$i].value -ge $SubstringFindObject)
                                {
                                    $IncidentsDataGridView.Rows[$j].visible=$false
                                }else{
                                    $IncidentsDataGridView.Rows[$j].visible=$True
                                    $VisibleGrid+=$IncidentsDataGridView.Rows[$j].DataBoundItem
                                }
                            }
                        }
                }
                
                break
            }
        }
        Write-Host $PropertyFindObject
        Write-Host $SubstringFindObject
    }
    if($SubstringFindObject.length -le 4)
    {
        for($j=0;$j -lt $IncidentsDataGridView.RowCount; $j++)
        {
            $IncidentsDataGridView.Rows[$j].visible=$true
            $VisibleGrid+=$IncidentsDataGridView.Rows[$j].DataBoundItem
        }
    }

    <#
    $TmpStr=$comboBox3Station.text
    $comboBox3Station.Items.Clear()
    $comboBox3Station.Items.AddRange(($VisibleGrid|Group-Object "Станция").Name)
    $comboBox3Station.text = $TmpStr

    $TmpStr=$comboBox3Point.text
    $comboBox3Point.Items.Clear()
    $comboBox3Point.Items.AddRange(($VisibleGrid|Group-Object "Узел").Name)
    $comboBox3Point.text = $TmpStr
    #>
}

function Get-SmenaReportFile(){
    $SmenaReport = Get-Content (dir "$env:USERPROFILE\Documents\SPD_Reports\*-report.txt" | Sort-Object LastWriteTime | select -Last 1).fullname
    foreach($StrSmenaReport in $SmenaReport)
    {
        $textBox2.Text += $StrSmenaReport + "`r`n"
    }
}

function Get-DayReportFile(){
    $SmenaReport = Get-Content (dir "$env:USERPROFILE\Documents\SPD_Reports\*-Fullreport.txt" | Sort-Object LastWriteTime | select -Last 1).fullname
    foreach($StrSmenaReport in $SmenaReport)
    {
        $textBox4.Text += $StrSmenaReport + "`r`n"
    }
}

function Button_Click(){
    $textBox2.Text = ""
    $textBox4.Text = ""
    if($textBoxAccessINC.Text -match "ИНЦ[00-99]")
    {
        get-SmenaReport -NumbersIncidensFromPreviousSmena $textBox1.Text.Replace(" ","") -NumbersIncidentsSoonComplite $textBoxAccessINC.Text.Replace(" ","") -AddDublicatePoint $DublicateIncidensCheckBox.Checked
    }else{
        get-SmenaReport -NumbersIncidensFromPreviousSmena $textBox1.Text.Replace(" ","") -AddDublicatePoint $DublicateIncidensCheckBox.Checked
    }

    if($textBox1.Text -like ""){$textBox1.Text = "Введите номера инцидентов переданных по смене через запятую"}
    if($textBoxAccessINC.Text -like ""){$textBoxAccessINC.Text = "Введите номера инцидентов, которые закроются на момент передачи смены"}

    Get-SmenaReportFile
    Get-DayReportFile
    Fill-IncidentsDataGridView
}


function start-form()
{
    Add-Type -AssemblyName System.Windows.Forms

    # Главная Форма
    $DublicateIncidensCheckBox = new-object System.Windows.Forms.CheckBox
    $label1 = new-object System.Windows.Forms.Label
    $button1 = new-object System.Windows.Forms.Button
    $textBox1 = new-object System.Windows.Forms.TextBox
    $textBoxAccessINC = new-object System.Windows.Forms.TextBox
    # Набор закладок
    $tabControl1 = New-Object System.Windows.Forms.TabControl
        # Закладка №1
        $SmenaReport = new-object System.Windows.Forms.TabPage
            $textBox2 = new-object System.Windows.Forms.TextBox
        # Закладка №2
        $DayReport = new-object System.Windows.Forms.TabPage
            $textBox4 = new-object System.Windows.Forms.TextBox
        # Закладка №3
        $IncidentsTab = new-object System.Windows.Forms.TabPage
            $comboBox3Zone = new-object System.Windows.Forms.ComboBox
            $labe3lZone = new-object System.Windows.Forms.Label
            $label3IncNumber = new-object System.Windows.Forms.Label
            $textBox3IncNumber = new-object System.Windows.Forms.TextBox
            $labelShort3Information = new-object System.Windows.Forms.Label
            $textBox3ShortInformation = new-object System.Windows.Forms.TextBox
            $label3Solved = new-object System.Windows.Forms.Label
            $textBox3Solved = new-object System.Windows.Forms.TextBox
            $dateTimePicker3TimeStart = new-object System.Windows.Forms.DateTimePicker
            $label3Station = new-object System.Windows.Forms.Label
            $comboBox3Station = new-object System.Windows.Forms.ComboBox
            $comboBox3Point = new-object System.Windows.Forms.ComboBox
            $label3Point = new-object System.Windows.Forms.Label
            $label3TimeStart = new-object System.Windows.Forms.Label
            $label3TimeEnd = new-object System.Windows.Forms.Label
            $dateTimePicker3TimeEnd = new-object System.Windows.Forms.DateTimePicker
            $label3IP = new-object System.Windows.Forms.Label
            $textBox3IP = new-object System.Windows.Forms.TextBox
            $IncidentsDataGridView = new-object System.Windows.Forms.DataGridView
        # Закладка №4
        $tabPage_ElectronicNotoficators = new-object System.Windows.Forms.TabPage
            $button4_ECHZONE = new-object System.Windows.Forms.Button
            $textBox4_ECHZone = new-object System.Windows.Forms.TextBox
            $label4_ECHZone = new-object System.Windows.Forms.Label

    ## 
    ## textBox1
    ## 
    $textBox1.Anchor = (((([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left) -bor `
    [System.Windows.Forms.AnchorStyles]::Right)))
    $textBox1.Location = new-object System.Drawing.Point(12, 20)
    $textBox1.Name = "textBox1"
    $textBox1.RightToLeft = [System.Windows.Forms.RightToLeft]::No
    $textBox1.Size = new-object System.Drawing.Size(672, 20)
    $textBox1.TabIndex = 1
    $textBox1.Text = "Введите номера инцидентов переданных по смене через запятую"
            ## 
            ## tabControl1
            ## 
            $tabControl1.Anchor = ((((([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom) -bor `
    [System.Windows.Forms.AnchorStyles]::Left)  -bor [System.Windows.Forms.AnchorStyles]::Right)))
            $tabControl1.Controls.Add($SmenaReport)
            $tabControl1.Controls.Add($DayReport)
            $tabControl1.Controls.Add($IncidentsTab)
            $tabControl1.Controls.Add($tabPage_ElectronicNotoficators);
            $tabControl1.Location = new-object System.Drawing.Point(15, 75)
            $tabControl1.Name = "tabControl1"
            $tabControl1.SelectedIndex = 0
            $tabControl1.Size = new-object System.Drawing.Size(773, 751)
            $tabControl1.TabIndex = 5
            ## 
            ## SmenaReport
            ## 
            $SmenaReport.Controls.Add($textBox2)
            $SmenaReport.Location = new-object System.Drawing.Point(4, 22)
            $SmenaReport.Name = "SmenaReport"
            $SmenaReport.Padding = new-object System.Windows.Forms.Padding(3)
            $SmenaReport.Size = new-object System.Drawing.Size(765, 725)
            $SmenaReport.TabIndex = 0
            $SmenaReport.Text = "SmenaReport"
            $SmenaReport.UseVisualStyleBackColor = $true
            ## 
            ## DayReport
            ## 
            $DayReport.Controls.Add($textBox4)
            $DayReport.Location = new-object System.Drawing.Point(4, 22)
            $DayReport.Name = "DayReport"
            $DayReport.Padding = new-object System.Windows.Forms.Padding(3)
            $DayReport.Size = new-object System.Drawing.Size(765, 725)
            $DayReport.TabIndex = 1
            $DayReport.Text = "DayReport"
            $DayReport.UseVisualStyleBackColor = $true
            ## 
            ## IncidentsTab
            ## 
            $IncidentsTab.Controls.Add($IncidentsDataGridView)
            $IncidentsTab.Controls.Add($dateTimePicker3TimeEnd)
            $IncidentsTab.Controls.Add($dateTimePicker3TimeStart)
            $IncidentsTab.Controls.Add($textBox3IP)
            $IncidentsTab.Controls.Add($label3IP)
            $IncidentsTab.Controls.Add($textBox3Solved)
            $IncidentsTab.Controls.Add($label3Solved)
            $IncidentsTab.Controls.Add($textBox3ShortInformation)
            $IncidentsTab.Controls.Add($labelShort3Information)
            $IncidentsTab.Controls.Add($label3Point)
            $IncidentsTab.Controls.Add($label3Station)
            $IncidentsTab.Controls.Add($textBox3IncNumber)
            $IncidentsTab.Controls.Add($label3IncNumber)
            $IncidentsTab.Controls.Add($label3TimeEnd)
            $IncidentsTab.Controls.Add($label3TimeStart)
            $IncidentsTab.Controls.Add($labe3lZone)
            $IncidentsTab.Controls.Add($comboBox3Point)
            $IncidentsTab.Controls.Add($comboBox3Station)
            $IncidentsTab.Controls.Add($comboBox3Zone)
            $IncidentsTab.Location = new-object System.Drawing.Point(4, 22)
            $IncidentsTab.Name = "IncidentsTab"
            $IncidentsTab.Padding = New-Object System.Windows.Forms.Padding(3)
            $IncidentsTab.Size = New-Object System.Drawing.Size(765, 725)
            $IncidentsTab.TabIndex = 2
            $IncidentsTab.Text = "Инциденты"
            $IncidentsTab.UseVisualStyleBackColor = $true
                ## 
                ## comboBox3Zone
                ## 
                $comboBox3Zone.FormattingEnabled = $true
                $comboBox3Zone.Location = new-object System.Drawing.Point(80, 3)
                $comboBox3Zone.Name = "comboBox3Zone"
                $comboBox3Zone.Size = new-object System.Drawing.Size(121, 21)
                $comboBox3Zone.TabIndex = 1
                #$comboBox3Zone.Items.
                ## 
                ## labe3lZone
                ## 
                $labe3lZone.AutoSize = $true
                $labe3lZone.Location = new-object System.Drawing.Point(6, 6)
                $labe3lZone.Name = "labe3lZone"
                $labe3lZone.Size = new-object System.Drawing.Size(32, 13)
                $labe3lZone.TabIndex = 2
                $labe3lZone.Text = "Зона"
                ## 
                ## label3IncNumber
                ## 
                $label3IncNumber.AutoSize = $true
                $label3IncNumber.Location = new-object System.Drawing.Point(207, 54)
                $label3IncNumber.Name = "label3IncNumber"
                $label3IncNumber.Size = new-object System.Drawing.Size(68, 13)
                $label3IncNumber.TabIndex = 3
                $label3IncNumber.Text = "Номер ИНЦ"
                ## 
                ## textBox3IncNumber
                ## 
                $textBox3IncNumber.Location = new-object System.Drawing.Point(296, 51)
                $textBox3IncNumber.Name = "textBox3IncNumber"
                $textBox3IncNumber.Size = new-object System.Drawing.Size(150, 20)
                $textBox3IncNumber.TabIndex = 4
                ## 
                ## labelShort3Information
                ## 
                $labelShort3Information.AutoSize = $true
                $labelShort3Information.Location = new-object System.Drawing.Point(452, 6)
                $labelShort3Information.Name = "labelShort3Information"
                $labelShort3Information.Size = new-object System.Drawing.Size(114, 13)
                $labelShort3Information.TabIndex = 3
                $labelShort3Information.Text = "Краткое содержание"
                ## 
                ## textBox3ShortInformation
                ## 
                $textBox3ShortInformation.Location = new-object System.Drawing.Point(569, 3)
                $textBox3ShortInformation.Name = "textBox3ShortInformation"
                $textBox3ShortInformation.Size = new-object System.Drawing.Size(190, 20)
                $textBox3ShortInformation.TabIndex = 4
                ## 
                ## label3Solved
                ## 
                $label3Solved.AutoSize = $true
                $label3Solved.Location = new-object System.Drawing.Point(452, 30)
                $label3Solved.Name = "label3Solved"
                $label3Solved.Size = new-object System.Drawing.Size(52, 13)
                $label3Solved.TabIndex = 3
                $label3Solved.Text = "Решение"
                ## 
                ## textBox3Solved
                ## 
                $textBox3Solved.Location = new-object System.Drawing.Point(568, 30)
                $textBox3Solved.Name = "textBox3Solved"
                $textBox3Solved.Size = new-object System.Drawing.Size(191, 20)
                $textBox3Solved.TabIndex = 4
                ## 
                ## dateTimePicker3TimeStart
                ## 
                $dateTimePicker3TimeStart.Location = new-object System.Drawing.Point(296, 3)
                $dateTimePicker3TimeStart.Name = "dateTimePicker3TimeStart"
                $dateTimePicker3TimeStart.Size = new-object System.Drawing.Size(150, 20)
                $dateTimePicker3TimeStart.TabIndex = 5
                ## 
                ## label3Station
                ## 
                $label3Station.AutoSize = $true
                $label3Station.Location = new-object System.Drawing.Point(6, 30)
                $label3Station.Name = "label3Station"
                $label3Station.Size = new-object System.Drawing.Size(49, 13)
                $label3Station.TabIndex = 3
                $label3Station.Text = "Станция"
                ## 
                ## comboBox3Station
                ## 
                $comboBox3Station.FormattingEnabled = $true

                $comboBox3Station.Location = new-object System.Drawing.Point(80, 27)
                $comboBox3Station.Name = "comboBox3Station"
                $comboBox3Station.Size = new-object System.Drawing.Size(121, 21)
                $comboBox3Station.TabIndex = 1
                ## 
                ## comboBox3Point
                ## 
                $comboBox3Point.FormattingEnabled = $true
                $comboBox3Point.Location = new-object System.Drawing.Point(80, 51)
                $comboBox3Point.Name = "comboBox3Point"
                $comboBox3Point.Size = new-object System.Drawing.Size(121, 21)
                $comboBox3Point.TabIndex = 1
                ## 
                ## label3Point
                ## 
                $label3Point.AutoSize = $true
                $label3Point.Location = new-object System.Drawing.Point(6, 54)
                $label3Point.Name = "label3Point"
                $label3Point.Size = new-object System.Drawing.Size(33, 13)
                $label3Point.TabIndex = 3
                $label3Point.Text = "Узел"
                ## 
                ## label3TimeStart
                ## 
                $label3TimeStart.AutoSize = $true
                $label3TimeStart.Location = new-object System.Drawing.Point(207, 6)
                $label3TimeStart.Name = "label3TimeStart"
                $label3TimeStart.Size = new-object System.Drawing.Size(59, 13)
                $label3TimeStart.TabIndex = 2
                $label3TimeStart.Text = "Начиная с"
                ## 
                ## label3TimeEnd
                ## 
                $label3TimeEnd.AutoSize = $true
                $label3TimeEnd.Location = new-object System.Drawing.Point(207, 30)
                $label3TimeEnd.Name = "label3TimeEnd"
                $label3TimeEnd.Size = new-object System.Drawing.Size(67, 13)
                $label3TimeEnd.TabIndex = 2
                $label3TimeEnd.Text = "Заканчивая"
                ## 
                ## dateTimePicker3TimeEnd
                ## 
                $dateTimePicker3TimeEnd.Location = new-object System.Drawing.Point(296, 27)
                $dateTimePicker3TimeEnd.Name = "dateTimePicker3TimeEnd"
                $dateTimePicker3TimeEnd.Size = new-object System.Drawing.Size(150, 20)
                $dateTimePicker3TimeEnd.TabIndex = 5
                ## 
                ## label3IP
                ## 
                $label3IP.AutoSize = $true
                $label3IP.Location = new-object System.Drawing.Point(452, 54)
                $label3IP.Name = "label3IP"
                $label3IP.Size = new-object System.Drawing.Size(17, 13)
                $label3IP.TabIndex = 3
                $label3IP.Text = "IP"
                ## 
                ## textBox3IP
                ## 
                $textBox3IP.Location = new-object System.Drawing.Point(568, 54)
                $textBox3IP.Name = "textBox3IP"
                $textBox3IP.Size = new-object System.Drawing.Size(191, 20)
                $textBox3IP.TabIndex = 4
                ## 
                ## IncidentsDataGridView
                ## 
                $IncidentsDataGridView.Anchor = ((((([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom) -bor `
                [System.Windows.Forms.AnchorStyles]::Left)  -bor [System.Windows.Forms.AnchorStyles]::Right)))
                $IncidentsDataGridView.AllowUserToOrderColumns=$true
                $IncidentsDataGridView.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
                $IncidentsDataGridView.Location = new-object System.Drawing.Point(0, 80)
                $IncidentsDataGridView.Name = "IncidentsDataGridView"
                $IncidentsDataGridView.Size = new-object System.Drawing.Size(763, 643)
                $IncidentsDataGridView.TabIndex = 0
            ## 
            ## tabPage_ElectronicNotoficators
            ## 
			$tabPage_ElectronicNotoficators.Controls.Add($label4_ECHZone);
            $tabPage_ElectronicNotoficators.Controls.Add($textBox4_ECHZone);
            $tabPage_ElectronicNotoficators.Controls.Add($button4_ECHZONE);
            $tabPage_ElectronicNotoficators.Location = new-object System.Drawing.Point(4, 22);
            $tabPage_ElectronicNotoficators.Name = "tabPage_ElectronicNotoficators";
            $tabPage_ElectronicNotoficators.Padding = new-object System.Windows.Forms.Padding(3);
            $tabPage_ElectronicNotoficators.Size = new-object System.Drawing.Size(792, 725);
            $tabPage_ElectronicNotoficators.TabIndex = 3;
            $tabPage_ElectronicNotoficators.Text = "ЭЧ Уведомления";
            $tabPage_ElectronicNotoficators.UseVisualStyleBackColor = $true;
                ## 
                ## button4_ECHZONE
                ## 
                $button4_ECHZONE.Location = new-object System.Drawing.Point(685, 6); #700
                $button4_ECHZONE.Name = "button4_ECHZONE";
                $button4_ECHZONE.Size = new-object System.Drawing.Size(75, 23);
                $button4_ECHZONE.TabIndex = 0;
                $button4_ECHZONE.Text = "Подготовить";
                $button4_ECHZONE.UseVisualStyleBackColor = $true;
                    ### button4_ECHZONE Action
                    $button4_ECHZONE.Add_Click({$Form1.Hide();
                            button4_ECHZONE_Click;
                            $Form1.Show()})

                ## 
                ## textBox4_ECHZone
                ## 
                $textBox4_ECHZone.Dock = [System.Windows.Forms.DockStyle]::Bottom;
                $textBox4_ECHZone.Location = new-object System.Drawing.Point(3, 35);
                $textBox4_ECHZone.Multiline = $true;
                $textBox4_ECHZone.Name = "textBox4_ECHZone";
                $textBox4_ECHZone.Size = new-object System.Drawing.Size(750, 687);
                $textBox4_ECHZone.TabIndex = 1;
                ## 
                ## label4_ECHZone
                ## 
                $label4_ECHZone.AutoSize = $true;
                $label4_ECHZone.Location = new-object System.Drawing.Point(3, 16);
                $label4_ECHZone.Name = "label4_ECHZone";
                $label4_ECHZone.Size = new-object System.Drawing.Size(198, 13);
                $label4_ECHZone.TabIndex = 2;
                $label4_ECHZone.Text = "Вставте текст заявки одной зоны ЭЧ";
    ## 
    ## textBox2
    ## 
    $textBox2.Anchor = ((((([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom) -bor `
    [System.Windows.Forms.AnchorStyles]::Left)  -bor [System.Windows.Forms.AnchorStyles]::Right)))
    #$textBox2.FormattingEnabled = $True
    $textBox2.Location = new-object System.Drawing.Point(0, 0)
    $textBox2.Multiline = $True
    $textBox2.Name = "textBox2"
    $textBox2.Size = new-object System.Drawing.Size(765, 719)
    $textBox2.ScrollBars = "Vertical"
    $textBox2.ReadOnly = $True
    $textBox2.TabIndex = 4
    ## 
    ## textBox4
    ## 
    $textBox4.Anchor = ((((([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom) -bor `
    [System.Windows.Forms.AnchorStyles]::Left)  -bor [System.Windows.Forms.AnchorStyles]::Right)))
    #$textBox2.FormattingEnabled = $True
    $textBox4.Location = new-object System.Drawing.Point(0, 0)
    $textBox4.Multiline = $True
    $textBox4.Name = "textBox2"
    $textBox4.Size = new-object System.Drawing.Size(765, 719)
    $textBox4.ScrollBars = "Vertical"
    $textBox4.ReadOnly = $True
    $textBox4.TabIndex = 4
    ##
	## textBoxAccessINC
    ## 
    $textBoxAccessINC.Anchor = (((([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left) -bor `
    [System.Windows.Forms.AnchorStyles]::Right)))
    $textBoxAccessINC.Location = new-object System.Drawing.Point(12, 49)
    $textBoxAccessINC.Name = "textBoxAccessINC"
    $textBoxAccessINC.RightToLeft = [System.Windows.Forms.RightToLeft]::No
    $textBoxAccessINC.Size = new-object System.Drawing.Size(672, 20)
    $textBoxAccessINC.TabIndex = 5
    $textBoxAccessINC.Text = "Введите номера инцидентов, которые закроются на момент передачи смены"
    if(Test-Path "$env:USERPROFILE\Documents\SPD_Reports\NumbersIncidentsSoonComplite.txt"){
        $textBoxAccessINC.Text = Get-Content "$env:USERPROFILE\Documents\SPD_Reports\NumbersIncidentsSoonComplite.txt"
    }



    ## 
    ## button1
    ## 
    $button1.Anchor = ((([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)))
    $button1.Location = new-object System.Drawing.Point(706, 32)
    $button1.Name = "button1"
    $button1.Size = new-object System.Drawing.Size(82, 23)
    $button1.TabIndex = 2
    $button1.Text = "Подготовить"
    $button1.UseVisualStyleBackColor = $True
    ## 
    ## label1
    ## 
    $label1.AutoSize = $True
    $label1.Location = new-object System.Drawing.Point(2, 3)
    $label1.Name = "label1"
    $label1.Size = new-object System.Drawing.Size(62, 13)
    $label1.TabIndex = 3
    $label1.Text = "Отчет от ..."
    ##
    ## DublicateIncidensCheckBox
    ##
    $DublicateIncidensCheckBox.Text = 'Добавить зависимости'
    $DublicateIncidensCheckBox.AutoSize = $true
    $DublicateIncidensCheckBox.Checked = $true
    $DublicateIncidensCheckBox.Location  = New-Object System.Drawing.Point(120,3)
    ## 
    ## Form1
    ## 
    #$AutoScaleDimensions = new-object System.Drawing.SizeF(6F, 13F)
    #$AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    [System.Windows.Forms.Form]$form1 = New-Object System.Windows.Forms.Form;
    $form1.Width = 500;
    $form1.Height = 500;
    $form1.ClientSize = new-object System.Drawing.Size(800, 838)
    $form1.Controls.Add($label1)
    $form1.Controls.Add($textBox1)
    $form1.Controls.Add($TabControl1)#$form1.Controls.Add($textBox2)
    $form1.Controls.Add($textBoxAccessINC)
    $form1.Controls.Add($button1)
    $form1.Controls.Add($DublicateIncidensCheckBox)
    $form1.Name = "Form1"
    $form1.Text = "Подготовка отчета смены"
    #$ResumeLayout($False)
    #$PerformLayout

    #actions
    $button1.Add_Click({Button_Click})
    Button_Click
    $textBox1.Add_KeyDown({
        if ($_.KeyCode -eq "Enter") {
            Button_Click
        }
    })
    $textBoxAccessINC.Add_KeyDown({
        if ($_.KeyCode -eq "Enter") {
            Button_Click
        }
    })

    # Старт поиска при изменении текста в полях филтрации
    $textBox3IncNumber.Add_TextChanged({find-incidents -PropertyFindObject $textBox3IncNumber.Name -SubstringFindObject $textBox3IncNumber.Text})
    $comboBox3Zone.Add_TextChanged({find-incidents -PropertyFindObject $comboBox3Zone.Name -SubstringFindObject $comboBox3Zone.Text})
    $comboBox3Point.Add_TextChanged({find-incidents -PropertyFindObject $comboBox3Point.Name -SubstringFindObject $comboBox3Point.Text})
    $comboBox3Station.Add_TextChanged({find-incidents -PropertyFindObject $comboBox3Station.Name -SubstringFindObject $comboBox3Station.Text})
    $textBox3Solved.Add_TextChanged({find-incidents -PropertyFindObject $textBox3Solved.Name -SubstringFindObject $textBox3Solved.Text})
    $textBox3IP.Add_TextChanged({find-incidents -PropertyFindObject $textBox3IP.Name -SubstringFindObject $textBox3IP.Text})

    #### !!!!!!!!!!!!!! Поиск по дате не работает , неверное приведение типов "Не удается преобразовать значение "25 сентября 2019 г." в тип "System.DateTime"
    $dateTimePicker3TimeStart.Add_TextChanged({find-incidents -PropertyFindObject $dateTimePicker3TimeStart.Name -SubstringFindObject $dateTimePicker3TimeStart.Value})
    $dateTimePicker3TimeEnd.Add_TextChanged({find-incidents -PropertyFindObject $dateTimePicker3TimeEnd.Name -SubstringFindObject $dateTimePicker3TimeEnd.Value})
    

    

    #[System.Windows.MessageBox]::Show("text changed")
    $IncidentsDataGridView.Add_CellMouseDoubleClick({grid_Click})

    $form1.ShowDialog() | Out-Null;
    
    write-errorlog
}

function write-errorlog()
{
    $ErrDir = "$env:USERPROFILE\Documents\SPD_Reports\ErrDir"
    $ErrFile = "StartSmenaRep_ErrorLog.log"

    if(!(Test-Path $ErrDir)){New-Item -ItemType Directory $ErrDir}
    if($Error)
    {
        echo "_______________" >> "$ErrDir\$ErrFile"
        Get-Date >> "$ErrDir\$ErrFile"
        $Error >> "$ErrDir\$ErrFile"
    }
}

start-form
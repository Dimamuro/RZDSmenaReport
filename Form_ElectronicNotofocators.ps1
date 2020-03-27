#param($FormObject = Get-Content "$env:USERPROFILE\Documents\SPD_Reports\Data\ECH_Notificators.txt")
$Error.Clear()

###    Function for generate text to notificators cells
function Out-ElectricNotificatos($InPutTextElectronicNotificators)
{
    function Get-ObjectPropery
    {
        param($TimeStart,
                $TimeEnd,
                $Station,
                $EK)

        $Object = New-Object psobject
        $Object|Add-Member -MemberType NoteProperty -Name TimeStart -Value $TimeStart -Force
        $Object|Add-Member -MemberType NoteProperty -Name TimeEnd -Value $TimeEnd -Force
        $Object|Add-Member -MemberType NoteProperty -Name Station -Value $Station -Force
        $Object|Add-Member -MemberType NoteProperty -Name EK -Value $EK -Force
        #$Object|Add-Member -MemberType NoteProperty -Name DateInaccssibility -Value $DateInaccssibility -Force
        #$Object|Add-Member -MemberType NoteProperty -Name ECHZone -Value $ECHZone -Force

        return $Object
    }

    $Glossary = Get-Content "C:\windows\System32\WindowsPowerShell\v1.0\Modules\GetSmenaReport\Glossary.txt"

    $Stations = @()
    foreach($tmpstr in $InPutTextElectronicNotificators)
    {
        $LengthStr = $tmpstr.Length
        for($i=$LengthStr-1; $i -ge 0; $i--)
        {
            if(!($tmpstr[$i] -match "\W"))
            {
                $tmpstr = $tmpstr.Substring(0,$i+1)
                $i = -1
            }
        }
        ### Дата начала простоя
        if($tmpstr -match "\d\d\.\d\d\.\d\d")
        {
            $DateInaccssibility = $tmpstr.Split(" ")|where{$_ -match "\d\d\.\d\d\.\d\d"}
            $StartActive = (get-date $DateInaccssibility -Format "dd'/'MM'/'yy HH:mm:ss").ToString()
            $EndActive = (get-date (get-date $DateInaccssibility).AddDays(1) -Format "dd'/'MM'/'yy HH:mm:ss").ToString()

            $textBox_StartActive.Text = $StartActive
            $textBox_EndActive.Text = $EndActive
        } 
        
        ###  Переменная для время овончания простоя
        $EndActiveTMP = get-date (get-date $DateInaccssibility).AddDays(1)

        ### Зона ЭЧ
        if($tmpstr -match "ЭЧ-")
        {
            $IndexOf = $tmpstr.IndexOf("ЭЧ-")
            $ECHZone = $tmpstr.Substring($IndexOf,4)
        }

        ### Станции из Словаря
        if($tmpstr.Replace(" ","") -match "^-")
        {
            if($tmpstr -match "ё"){$tmpstr = $tmpstr.replace("ё","е")}
            foreach($GlossaryStation in $Glossary)
            {
                $GlossaryStation = $GlossaryStation.replace("-"," ")
                if($tmpstr -match $GlossaryStation)
                {
                    if($tmpstr -match "\d:\d\d-")
                    {
                        $TimeRange = ($tmpstr.Split(" ")|where{$_ -match "\d:\d\d-"})
                        [datetime]$TimeStart = (get-date $DateInaccssibility) + ($TimeRange.Split("-")[0])+("04:00")                    
                        [datetime]$TimeEnd = (get-date $DateInaccssibility) + ($TimeRange.Split("-")[-1])+("04:00")
                    }else{
                        [datetime]$TimeStart = (get-date $DateInaccssibility) + [timespan]("08:00")
                        [datetime]$TimeEnd = (get-date $DateInaccssibility) + [timespan]("17:00")
                    }

                    if($TimeStart -gt $TimeEnd){$TimeEnd = $TimeEnd.AddDays(1)}

                    if($EndActiveTMP -lt $TimeEnd)
                    {
                        $EndActiveTMP = $TimeEnd
                        $textBox_EndActive.Text = (get-date ($EndActiveTMP) -Format "dd'/'MM'/'yy HH:mm:ss").ToString()
                    }

                    $TimeStart_TypeStr = get-date ($TimeStart) -Format "dd'/'MM'/'yy HH:mm:ss"
                    $TimeEnd_TypeStr = get-date ($TimeEnd) -Format "dd'/'MM'/'yy HH:mm:ss"
                    $Stations += Get-ObjectPropery -TimeStart $TimeStart_TypeStr -TimeEnd $TimeEnd_TypeStr `
                            -Station $GlossaryStation -EK "`*$GlossaryStation`*КРАСН".Replace(" ","")
                }
            }      
        } 
    }

    $dataGridView_Inaccessibility.DataSource = [System.Collections.ArrayList]$Stations
    $ColStartDateInaccessibility = $dataGridView_Inaccessibility.Columns[0]
    $ColEndDateInaccessibility = $dataGridView_Inaccessibility.Columns[1]
    $ColStationInaccessibility = $dataGridView_Inaccessibility.Columns[2]
    $ColEKText = $dataGridView_Inaccessibility.Columns[3]

    <#Скрываю левые запросы
    foreach($station in $dataGridView_Inaccessibility.Rows)
    {
        if($station[3] -match "Ирба" )
    }
    #>

    #Скрываю колонку с наименованием станции
    $ColStationInaccessibility.visible=$false


    ## ColStartDateInaccessibility
    ## 
    $ColStartDateInaccessibility.HeaderText = "Начало простоя";
    $ColStartDateInaccessibility.Name = "ColStartDateInaccessibility";
    $ColStartDateInaccessibility.Width = 190;
    ## 
    ## ColEndDateInaccessibility
    ## 
    $ColEndDateInaccessibility.HeaderText = "Оконание простоя";
    $ColEndDateInaccessibility.Name = "ColEndDateInaccessibility";
    $ColEndDateInaccessibility.Width = 190;
    ## 
    ## ColEKText
    ## 
    $ColEKText.HeaderText = "ЭК";
    $ColEKText.Name = "ColEKText";
    $ColEKText.Width = 385;



    $textBox_Description.Lines += "Плановые работы $ECHZone cо снятием напряжения на $DateInaccssibility`:"
    $textBox_Description.Lines += ""
    foreach($tmpstr In ($Stations|Group-Object Station).name)
    {
        $textBox_Description.text += $tmpstr + ", "
    }
}

function start-form()
{
    param($InPutTextElectronicNotificators)

    Add-Type -AssemblyName System.Windows.Forms

    $label1 = new-object System.Windows.Forms.Label
    $label_Level = new-object System.Windows.Forms.Label
    $label_Zone = new-object System.Windows.Forms.Label
    $label_AccessFor = new-object System.Windows.Forms.Label
    $label_Priority = new-object System.Windows.Forms.Label
    $label_StartActive = new-object System.Windows.Forms.Label
    $label_EndActive = new-object System.Windows.Forms.Label
    $label_Inaccessibility = new-object System.Windows.Forms.Label
    $label_EK = new-object System.Windows.Forms.Label
    $label_Category = new-object System.Windows.Forms.Label
    $label_System = new-object System.Windows.Forms.Label
    $label_Information = new-object System.Windows.Forms.Label
    $label_Description = new-object System.Windows.Forms.Label
    $textBox_Level = new-object System.Windows.Forms.TextBox
    $textBox_Zone = new-object System.Windows.Forms.TextBox
    $textBox_AccessFor = new-object System.Windows.Forms.TextBox
    $textBox_Priority = new-object System.Windows.Forms.TextBox
    $textBox_StartActive = new-object System.Windows.Forms.TextBox
    $textBox_EndActive = new-object System.Windows.Forms.TextBox
    $textBox_EK = new-object System.Windows.Forms.TextBox
    $textBox_Category = new-object System.Windows.Forms.TextBox
    $textBox_System = new-object System.Windows.Forms.TextBox
    $textBox_Information = new-object System.Windows.Forms.TextBox
    $textBox_Description = new-object System.Windows.Forms.TextBox
    $dataGridView_Inaccessibility = new-object System.Windows.Forms.DataGridView
    $ColStartDateInaccessibility = new-object System.Windows.Forms.DataGridViewTextBoxColumn
    $ColEndDateInaccessibility = new-object System.Windows.Forms.DataGridViewTextBoxColumn
    $ColEKText = new-object System.Windows.Forms.DataGridViewTextBoxColumn
    
    
    #((System.ComponentModel.ISupportInitialize)($dataGridView_Inaccessibility)).BeginInit
    #$SuspendLayout
    
    
    ## 
    ## label1
    ## 
    $label1.AutoSize = $true;
    $label1.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 12, [System.Drawing.FontStyle]::Bold);
    $label1.Location = new-object System.Drawing.Point(12, 9);
    $label1.Name = "label1";
    $label1.Size = new-object System.Drawing.Size(125, 20);
    $label1.TabIndex = 0;
    $label1.Text = "Уведомление";
    ## 
    ## label_Level
    ## 
    $label_Level.AutoSize = $true;
    $label_Level.Location = new-object System.Drawing.Point(15, 74);
    $label_Level.Name = "label_Level";
    $label_Level.Size = new-object System.Drawing.Size(54, 13);
    $label_Level.TabIndex = 1;
    $label_Level.Text = "Уровень:";
    ## 
    ## label_Zone
    ## 
    $label_Zone.AutoSize = $true;
    $label_Zone.Location = new-object System.Drawing.Point(15, 96);
    $label_Zone.Name = "label_Zone";
    $label_Zone.Size = new-object System.Drawing.Size(121, 13);
    $label_Zone.TabIndex = 1;
    $label_Zone.Text = "Зона ответственности";
    ## 
    ## label_AccessFor
    ## 
    $label_AccessFor.AutoSize = $true;
    $label_AccessFor.Location = new-object System.Drawing.Point(15, 119);
    $label_AccessFor.Name = "label_AccessFor";
    $label_AccessFor.Size = new-object System.Drawing.Size(77, 13);
    $label_AccessFor.TabIndex = 1;
    $label_AccessFor.Text = "Доступно для";
    ## 
    ## label_Priority
    ## 
    $label_Priority.AutoSize = $true;
    $label_Priority.Location = new-object System.Drawing.Point(15, 141);
    $label_Priority.Name = "label_Priority";
    $label_Priority.Size = new-object System.Drawing.Size(131, 13);
    $label_Priority.TabIndex = 1;
    $label_Priority.Text = "Приоритет отображения";
    ## 
    ## label_StartActive
    ## 
    $label_StartActive.AutoSize = $true;
    $label_StartActive.Location = new-object System.Drawing.Point(15, 171);
    $label_StartActive.Name = "label_StartActive";
    $label_StartActive.Size = new-object System.Drawing.Size(58, 13);
    $label_StartActive.TabIndex = 1;
    $label_StartActive.Text = "Активно с";
    ## 
    ## label_EndActive
    ## 
    $label_EndActive.AutoSize = $true;
    $label_EndActive.Location = new-object System.Drawing.Point(15, 191);
    $label_EndActive.Name = "label_EndActive";
    $label_EndActive.Size = new-object System.Drawing.Size(64, 13);
    $label_EndActive.TabIndex = 1;
    $label_EndActive.Text = "Активно до";
    ## 
    ## label_Inaccessibility
    ## 
    $label_Inaccessibility.AutoSize = $true;
    $label_Inaccessibility.Location = new-object System.Drawing.Point(15, 347);
    $label_Inaccessibility.Name = "label_Inaccessibility";
    $label_Inaccessibility.Size = new-object System.Drawing.Size(94, 13);
    $label_Inaccessibility.TabIndex = 1;
    $label_Inaccessibility.Text = "Список простоев";
    ## 
    ## label_EK
    ## 
    $label_EK.AutoSize = $true;
    $label_EK.Location = new-object System.Drawing.Point(398, 74);
    $label_EK.Name = "label_EK";
    $label_EK.Size = new-object System.Drawing.Size(21, 13);
    $label_EK.TabIndex = 1;
    $label_EK.Text = "ЭК";
    ## 
    ## label_Category
    ## 
    $label_Category.AutoSize = $true;
    $label_Category.Location = new-object System.Drawing.Point(398, 96);
    $label_Category.Name = "label_Category";
    $label_Category.Size = new-object System.Drawing.Size(60, 13);
    $label_Category.TabIndex = 1;
    $label_Category.Text = "Категория";
    ## 
    ## label_System
    ## 
    $label_System.AutoSize = $true;
    $label_System.Location = new-object System.Drawing.Point(398, 119);
    $label_System.Name = "label_System";
    $label_System.Size = new-object System.Drawing.Size(51, 13);
    $label_System.TabIndex = 1;
    $label_System.Text = "Система";
    ## 
    ## label_Information
    ## 
    $label_Information.AutoSize = $true;
    $label_Information.Location = new-object System.Drawing.Point(398, 150);
    $label_Information.Name = "label_Information";
    $label_Information.Size = new-object System.Drawing.Size(160, 13);
    $label_Information.TabIndex = 1;
    $label_Information.Text = "Дополнительная информация";
    ## 
    ## label_Description
    ## 
    $label_Description.AutoSize = $true;
    $label_Description.Location = new-object System.Drawing.Point(16, 222);
    $label_Description.Name = "label_Description";
    $label_Description.Size = new-object System.Drawing.Size(57, 13);
    $label_Description.TabIndex = 1;
    $label_Description.Text = "Описание";
    ## 
    ## textBox_Level
    ## 
    $textBox_Level.Location = new-object System.Drawing.Point(174, 71);
    $textBox_Level.Name = "textBox_Level";
    $textBox_Level.ReadOnly = $true;
    $textBox_Level.Size = new-object System.Drawing.Size(218, 20);
    $textBox_Level.TabIndex = 2;
    $textBox_Level.Text = "дорожный";
    ## 
    ## textBox_Zone
    ## 
    $textBox_Zone.Location = new-object System.Drawing.Point(174, 92);
    $textBox_Zone.Name = "textBox_Zone";
    $textBox_Zone.ReadOnly = $true;
    $textBox_Zone.Size = new-object System.Drawing.Size(218, 20);
    $textBox_Zone.TabIndex = 2;
    $textBox_Zone.Text = "88-КРАСН";
    ## 
    ## textBox_AccessFor
    ## 
    $textBox_AccessFor.Location = new-object System.Drawing.Point(174, 113);
    $textBox_AccessFor.Name = "textBox_AccessFor";
    $textBox_AccessFor.ReadOnly = $true;
    $textBox_AccessFor.Size = new-object System.Drawing.Size(218, 20);
    $textBox_AccessFor.TabIndex = 2;
    $textBox_AccessFor.Text = "все";
    ## 
    ## textBox_Priority
    ## 
    $textBox_Priority.Location = new-object System.Drawing.Point(174, 134);
    $textBox_Priority.Name = "textBox_Priority";
    $textBox_Priority.ReadOnly = $true;
    $textBox_Priority.Size = new-object System.Drawing.Size(218, 20);
    $textBox_Priority.TabIndex = 2;
    $textBox_Priority.Text = "3-Низкий";
    ## 
    ## textBox_StartActive
    ## 
    $textBox_StartActive.Location = new-object System.Drawing.Point(174, 167);
    $textBox_StartActive.Name = "textBox_StartActive";
    $textBox_StartActive.Size = new-object System.Drawing.Size(218, 20);
    $textBox_StartActive.TabIndex = 2;
    ## 
    ## textBox_EndActive
    ## 
    $textBox_EndActive.Location = new-object System.Drawing.Point(174, 188);
    $textBox_EndActive.Name = "textBox_EndActive";
    $textBox_EndActive.Size = new-object System.Drawing.Size(218, 20);
    $textBox_EndActive.TabIndex = 2;
    ## 
    ## textBox_EK
    ## 
    $textBox_EK.Location = new-object System.Drawing.Point(570, 73);
    $textBox_EK.Name = "textBox_EK";
    $textBox_EK.Size = new-object System.Drawing.Size(218, 20);
    $textBox_EK.TabIndex = 2;
    $textBox_EK.Text = "*ченить*КРАСН";
    ## 
    ## textBox_Category
    ## 
    $textBox_Category.Location = new-object System.Drawing.Point(570, 94);
    $textBox_Category.Name = "textBox_Category";
    $textBox_Category.Size = new-object System.Drawing.Size(218, 20);
    $textBox_Category.TabIndex = 2;
    $textBox_Category.Text = "плановые работы";
    ## 
    ## textBox_System
    ## 
    $textBox_System.Location = new-object System.Drawing.Point(570, 115);
    $textBox_System.Name = "textBox_System";
    $textBox_System.Size = new-object System.Drawing.Size(218, 20);
    $textBox_System.TabIndex = 2;
    $textBox_System.Text = "Электроэнергия";
    ## 
    ## textBox_Information
    ## 
    $textBox_Information.Location = new-object System.Drawing.Point(401, 171);
    $textBox_Information.Multiline = $true;
    $textBox_Information.Name = "textBox_Information";
    $textBox_Information.Size = new-object System.Drawing.Size(387, 162);
    $textBox_Information.TabIndex = 2;
    $textBox_Information.ScrollBars = "Vertical"
    foreach($tmpstr in $InPutTextElectronicNotificators)
    {
        $textBox_Information.Lines += $tmpstr
    }
    ## 
    ## textBox_Description
    ## 
    $textBox_Description.Location = new-object System.Drawing.Point(16, 238);
    $textBox_Description.Multiline = $true;
    $textBox_Description.Name = "textBox_Description";
    $textBox_Description.Size = new-object System.Drawing.Size(376, 95);
    $textBox_Description.TabIndex = 2;
    $textBox_Description.ScrollBars = "Vertical"
    ## 
    ## dataGridView_Inaccessibility
    ## 
    $dataGridView_Inaccessibility.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
    <#$dataGridView_Inaccessibility.ColumnCount=3
    $ColStartDateInaccessibility = $dataGridView_Inaccessibility.Columns[0]
    $ColEndDateInaccessibility = $dataGridView_Inaccessibility.Columns[1]
    $ColEKText = $dataGridView_Inaccessibility.Columns[2]
    #>
    $dataGridView_Inaccessibility.Location = new-object System.Drawing.Point(16, 363);
    $dataGridView_Inaccessibility.Name = "dataGridView_Inaccessibility";
    $dataGridView_Inaccessibility.RowHeadersVisible = $false;
    $dataGridView_Inaccessibility.Size = new-object System.Drawing.Size(772, 332);
    $dataGridView_Inaccessibility.TabIndex = 3;
    <## 
    ## ColStartDateInaccessibility
    ## 
    $ColStartDateInaccessibility.HeaderText = "Начало простоя";
    $ColStartDateInaccessibility.Name = "ColStartDateInaccessibility";
    $ColStartDateInaccessibility.Width = 190;
    ## 
    ## ColEndDateInaccessibility
    ## 
    $ColEndDateInaccessibility.HeaderText = "Оконание простоя";
    $ColEndDateInaccessibility.Name = "ColEndDateInaccessibility";
    $ColEndDateInaccessibility.Width = 190;
    ## 
    ## ColEKText
    ## 
    $ColEKText.HeaderText = "ЭК";
    $ColEKText.Name = "ColEKText";
    $ColEKText.Width = 385;
    ## 
    ## Form_ElectronicNotificators
    ##>
    [System.Windows.Forms.Form]$form_ElectroniNotificators = New-Object System.Windows.Forms.Form;
    $form_ElectroniNotificators.Width = 111
    $form_ElectroniNotificators.Height = 319
    $form_ElectroniNotificators.ClientSize = new-object System.Drawing.Size(800, 707)

    $form_ElectroniNotificators.Controls.Add($dataGridView_Inaccessibility);
    $form_ElectroniNotificators.Controls.Add($textBox_EndActive);
    $form_ElectroniNotificators.Controls.Add($textBox_StartActive);
    $form_ElectroniNotificators.Controls.Add($textBox_Priority);
    $form_ElectroniNotificators.Controls.Add($textBox_Description);
    $form_ElectroniNotificators.Controls.Add($textBox_Information);
    $form_ElectroniNotificators.Controls.Add($textBox_System);
    $form_ElectroniNotificators.Controls.Add($textBox_Category);
    $form_ElectroniNotificators.Controls.Add($textBox_AccessFor);
    $form_ElectroniNotificators.Controls.Add($textBox_EK);
    $form_ElectroniNotificators.Controls.Add($textBox_Zone);
    $form_ElectroniNotificators.Controls.Add($textBox_Level);
    $form_ElectroniNotificators.Controls.Add($label_Description);
    $form_ElectroniNotificators.Controls.Add($label_Information);
    $form_ElectroniNotificators.Controls.Add($label_System);
    $form_ElectroniNotificators.Controls.Add($label_Category);
    $form_ElectroniNotificators.Controls.Add($label_EK);
    $form_ElectroniNotificators.Controls.Add($label_Inaccessibility);
    $form_ElectroniNotificators.Controls.Add($label_EndActive);
    $form_ElectroniNotificators.Controls.Add($label_StartActive);
    $form_ElectroniNotificators.Controls.Add($label_Priority);
    $form_ElectroniNotificators.Controls.Add($label_AccessFor);
    $form_ElectroniNotificators.Controls.Add($label_Zone);
    $form_ElectroniNotificators.Controls.Add($label_Level);
    $form_ElectroniNotificators.Controls.Add($label1);
    $form_ElectroniNotificators.Name = "Form_ElectronicNotificators";
    $form_ElectroniNotificators.Text = "Генератор ЭЧ Уведомлений";

    Out-ElectricNotificatos $InPutTextElectronicNotificators

    $form_ElectroniNotificators.ShowDialog() | Out-Null;

    write-errorlog
}

function write-errorlog()
{
    $ErrDir = "$env:USERPROFILE\Documents\SPD_Reports\ErrDir"
    $ErrFile = "From_electronicNotificators_ErrorLog.log"

    if(!(Test-Path $ErrDir)){New-Item -ItemType Directory $ErrDir}
    if($Error)
    {
        echo "_______________" >> "$ErrDir\$ErrFile"
        Get-Date >> "$ErrDir\$ErrFile"
        $Error >> "$ErrDir\$ErrFile"
    }
}

$InPutFile = dir "$env:USERPROFILE\Documents\SPD_Reports\Data\ECH_Notificators.txt"
$InPutTextElectronicNotificators = Get-Content $InPutFile.FullName

start-form -InPutTextElectronicNotificators $InPutTextElectronicNotificators

$base_folder = Get-Location
$s_file = Get-ChildItem Settings.ps1

function get_copyname
{
    param ($name)
    if (Test-Path $name)
    {
        $ex_file = Get-ChildItem $name
        $num = (
            Get-ChildItem ("$($ex_file.BaseName)*") |% {
            $_.Name | Select-String '\((.+)\)'} |% {
            $_.Matches[0].Groups[1].Value} | 
            Measure-Object -Maximum
            ).Maximum + 1
        $name = Join-Path (Get-Location) ($ex_file.BaseName + " ($num)" + $ex_file.Extension)
    }
    else
    {
        $name = Join-Path (Get-Location) ($name)
    }
    return $name
}
function move_btw
{
    param([int]$from, [int]$to)
    $dist = $to - $from
    if ($dist -gt 0) {return "{RIGHT}"*(+$dist)}
    else {return "{LEFT}"*(-$dist)}
}
function message
{
    param([string]$cap, [string]$msg, [switch]$nopopup, [switch]$notitle)
    if ($cap.Length -gt 0) {Write-Host $cap}
    if ($msg.Length -gt 0) {Write-Host $msg}
    if (-not $notitle.IsPresent) {$host.ui.RawUI.WindowTitle = $cap}
    if (-not $nopopup.IsPresent) {$wshell.popup($msg, 0, $cap) > $null}
}
function activate_1c
{
    $id_1c = (Get-Process "1cv8" -ErrorAction Ignore |? {$_.SI -eq $SessionID}).id
    if ($id_1c -eq $null) { message "Должна быть запущена 1С"; return $False }
    elseif ($id_1c -is [System.Object[]]) { message "Должна быть запущена одна 1С"; return $False }
    else {$wshell.AppActivate($id_1c) > $null}
    $wshell.SendKeys("") > $null
    return $True
}
function load_settings
{
    $new_s = & $s_file
    #if ('request_range' -notin $new_s.Keys) {$new_s.request_range='A1:A1'}
    $new_s.clients = [ordered]@{}
    Get-ChildItem (Join-Path ($base_folder) ("*.csv")) |% {Import-Csv $_ -UseCulture -Encoding Default} |% {$new_s.clients[$_.Email] = $_.Search}
    $Global:s = $new_s
}
function read_lines
{
    if (-not $Excel.Visible) {message "Excel закрыт"; continue}
        
    [System.Collections.ArrayList] $lines = @()
    #$crds = $Excel.Selection.Address() -split ':' -replace ('^\$', '') -split '\$'
    #if ($crds.Count -eq 2) {$crds = $crds * 2}
    $message = ""
    $Excel.Selection.Rows() | ForEach-Object {
        if ($_.RowHeight -eq 0) { return }
        $code_text = $_.Cells(1).MergeArea(1,1).Text.Trim()
        $count_text = $_.Cells($Excel.Selection.Columns().Count).MergeArea(1,1).Text.Trim()
        if ($code_text -eq "" -or $count_text -eq "") {return}
        if ($code_text -notmatch "\d{13}" -and $count_text -match "\d{13}") {$code_text, $count_text = $count_text, $code_text}
        $code = [long] ($code_text -replace '\D', '')
        $count = $count_text -replace "[^0-9\.,]", '' -replace ',', '.'
        $line = @{
            code_text = $code_text
            count_text = $count_text
            code = $code
            count = $count
            in_pieces = $count_text -match "ш" #-or $Excel.ActiveSheet.Cells(3, 1).Interior.ColorIndex -eq 36
            replace = $null
            remove = $null
            recalc = $null
            warning = $null
            }
        if ($line.code -eq "" -or $line.count -eq "") { return }
        if ($Excel.Selection.Columns().Count -eq 1) {$line.count = $lines.count + 1}

        if ($s.replace_codes.ContainsKey($line.code))
        {
            $line.replace = $s.replace_codes[$line.code]
            $message += "Замена: $($line.code) -> $($line.replace[0]) [$($line.count_text)] $($line.replace[1])`n"
            $line.code = $s.replace_codes[$line.code][0]
        }
        if ($s.remove_codes.ContainsKey($line.code))
        {
            $line.remove = $s.remove_codes[$line.code]
            $message += "Удалено: $($line.code) [$($line.count_text)] $($line.remove)`n"
            return
        }
        if ($s.recalc_codes.ContainsKey($line.code) -and $Excel.Selection.Columns().Count -ne 1 -and -not $line.in_pieces)
        {
            $line.recalc = $s.recalc_codes[$line.code]
            $line.count = [int]([decimal]$line.count * [int]$line.recalc[0])
            $line.in_pieces = $True
            $message += "Пересчитано: $($line.code) [$($line.count_text) -> $($line.count) шт] $($line.recalc[1])`n"
        }
        if ($s.warning_codes.ContainsKey($line.code))
        {
            $line.warning = $s.warning_codes[$line.code]
            $message += "Внимание: $($line.code) [$($line.count_text)] $($line.warning)`n"
        }
        $lines.Add($line) > $null
    }
    return $lines, $message
}

load_settings

Set-Location $s.folder

$Excel = New-Object -ComObject Excel.Application
$Excel.DisplayAlerts = $false
$ApplicationWorkBook = $null
$wshell = New-Object -ComObject Wscript.Shell
$watcher = New-Object System.IO.FileSystemWatcher (Get-Location)
$SessionID = (Get-Process -PID $PID).SessionId

Register-ObjectEvent -InputObject $watcher -EventName Renamed
Register-ObjectEvent -InputObject $watcher -EventName Created

message "Приём заявок запущен:" "$($watcher.Path)"

activate_1c > $null

while ($True)
{
    $e = Wait-Event
    $e | Remove-Event
    $file_path = $e[0].SourceEventArgs.FullPath
    $file = Get-ChildItem -LiteralPath $file_path -ErrorAction Ignore
    if ($file -eq $null) { continue }
    if ($file.Extension -notin @('.xls', '.xlsx', '.eml')) { continue }

    load_settings

    if ($file.Extension -in @('.eml'))
    {
        $adoStream = New-Object -ComObject 'ADODB.Stream'
        $adoStream.Open()
        $adoStream.LoadFromFile($file.FullName)
        $cdoMessageObject = New-Object -ComObject 'CDO.Message'
        $cdoMessageObject.DataSource.OpenObject($adoStream, '_Stream')
        $email = ($cdoMessageObject.From | Select-String '<(.+)>').Matches[0].Groups[1].Value
        $client = $s.clients[$email]
        if ($client  -eq $null)
        {
            if ($email -match "m(\d+)dir@eurotorg\.by") {$client = $Matches[1]}
            elseif ($email -match "m(\d+)tov@eurotorg\.by") {$client = $Matches[1]}
            elseif ($email -match "dbr(\d+)dir@dobronom\.by") {$client = $Matches[1]}
            elseif ($email -match "dbr(\d+)tov@dobronom\.by") {$client = $Matches[1]}
        }
	
        if ($cdoMessageObject.Attachments.Count -ge 1)
        {
            $att_file_path = get_copyname $cdoMessageObject.Attachments[1].FileName
            $cdoMessageObject.Attachments[1].SaveToFile($att_file_path)
        }

        if ($client -eq $null)
        {
            Write-Host
            message "Не найдено:" "$email`n"
        }
	    else
	    {
            Write-Host
	        message "Найдено по почте: $email" "$client" -nopopup -notitle
	        if (-not (activate_1c)) {Get-Event | Remove-Event; continue}
            Set-Clipboard $client
	        $wshell.SendKeys("^{F3}+{F10}{DOWN}{DOWN}{DOWN} {TAB}")
	    }
        # SWAPPED V^
        if ($cdoMessageObject.Attachments.Count -gt 1)
        {
            message "Открыто 1 из $($cdoMessageObject.Attachments.Count) приложений" "Остальные приложения загружать вручную"
        }
	    elseif ($cdoMessageObject.HTMLBody -match "<table.*>")
	    {
	        $cdoMessageObject.HTMLBody | Set-Clipboard
            if ($ApplicationWorkBook -ne $null) {try {$ApplicationWorkBook.Close() > $null} catch {}}
	        $ApplicationWorkBook = $Excel.WorkBooks.Add()
	        $Excel.ActiveSheet.Paste()
            $file_name = get_copyname ($file.BaseName + ".xls")
	        $ApplicationWorkBook.SaveAs($file_name)
	        $ApplicationWorkBook.Close()
	    }
        elseif ($cdoMessageObject.Attachments.Count -eq 0)
        {
            message "Нет приложений"
        }
        
        #Get-Event | Remove-Event
    }
    if ($file.Extension -in @('.xls', '.xlsx') -and $file.name -ne $ApplicationWorkBook.Name)
    {
        if ($s.search_from_clipboard)
        {
	        $email = (Get-Clipboard) -join '' -replace '\s', ''
            $client = $s.clients[$email]
            if ($client  -eq $null)
            {
                if ($email -match "m(\d+)dir@eurotorg\.by") {$client = $Matches[1]}
                elseif ($email -match "m(\d+)tov@eurotorg\.by") {$client = $Matches[1]}
                elseif ($email -match "dbr(\d+)dir@dobronom\.by") {$client = $Matches[1]}
                elseif ($email -match "dbr(\d+)tov@dobronom\.by") {$client = $Matches[1]}
            }
            activate_1c > $null
	        if ($client -eq $null)
            {
                Write-Host
                message "Не найдено:" "$email`n"
            }
	        else
	        {
	            Write-Host
	            message "Найдено по почте: $email" "$client" -nopopup -notitle
	            if (-not (activate_1c)) {Get-Event | Remove-Event; continue}
                Set-Clipboard $client
	            $wshell.SendKeys("^{F3}+{F10}{DOWN}{DOWN}{DOWN} {TAB}")
	        }
        }

        if ($ApplicationWorkBook -ne $null) {try {$ApplicationWorkBook.Close() > $null} catch {}}
        $ApplicationWorkBook = $Excel.WorkBooks.Open($file_path)
        message '' "Открыт файл: $file_path" -nopopup -notitle

        $left_dow = $null
        $right_up = $null
        
        $light_yellow_head_row = ($Excel.ActiveSheet.Columns(1).Rows() | Select-Object -First 100 |? {$_.Interior.ColorIndex -eq 36} | Select-Object -First 1).EntireRow

 	    if ($light_yellow_head_row -ne $null)
	    {
		    $code_cell = $light_yellow_head_row.Find("Штрих-код")
		    $count_cell = $light_yellow_head_row.Find("Количество заказано")
            if ($code_cell -eq $null -or $count_cell -eq $null)
            {
                message '' 'Нет столбца "Штрих-код" или "Количество заказано"'
                $left_up = $light_yellow_head_row.Columns(1)
		        $left_down = $left_up.EntireColumn.Rows($left_up.EntireColumn.Rows.Count).End(-4162)
		        $right_up = $light_yellow_head_row.Columns(1).End(-4161)
            }
            else
            {
                $left_up = $code_cell
		        $left_down = $left_up.EntireColumn.Rows($left_up.EntireColumn.Rows.Count).End(-4162)
		        $right_up = $count_cell

		        $Excel.ActiveSheet.Range($count_cell.Offset(1, 0), $count_cell.End(-4121)) |% {
			        $_.Value = $_.Text + " шт"
                }
		    }
	    }
	    else
	    {
            $left_up = $Excel.Range($s.code_start)
		    $left_down = $left_up.EntireColumn.Rows($left_up.EntireColumn.Rows.Count).End(-4162)
		    $right_up = $Excel.Range($s.count_start)
	    }
        $Excel.ActiveSheet.Range($left_down, $right_up).Select() > $null 
	    $Excel.Selection.UnMerge() > $null
	    $Excel.ActiveSheet.Range($Excel.ActiveSheet.Cells(1, 1), $Excel.ActiveSheet.Range($left_down, $right_up)).Select() > $null
	    $Excel.ActiveWindow.Zoom = $True
	    $Excel.ActiveWindow.Zoom *= 0.98
	    $Excel.ActiveSheet.Range($left_down, $right_up).Select() > $null
	
        $ApplicationWorkBook.Saved = $true
        $Excel.Visible = $true

	    $wshell.AppActivate((Get-Process Excel | Where-Object {$_.MainWindowHandle -eq $Excel.Hwnd}).Id) > $null
        activate_1c > $null

	    $lines, $message = read_lines
	    if ($Excel.Selection.Count -gt 1 -and $lines.count -eq 0) {message "Пустая заявка" "В области по умолчанию нет строк заявки"}

        Get-Event | Remove-Event
        continue 
    }
    elseif ($file.Extension -in @('.xls', '.xlsx') -and $file.name -eq $ApplicationWorkBook.Name)
    {
        if ($Excel.Selection.Columns().Count -eq 1)
        {
            $is_count = $False
            $Excel.Selection.Cells | ForEach-Object {
                if ($_.Value() -match "\d+" -and $_.Value() -notmatch "\d{13}")
                {
                    $is_count = $True
                    $_.Value() = $_.Text + " шт"
                }
            }
            if ($is_count)
            {
                message "Места заменены на штуки" "" -nopopup -notitle
                Get-Event | Remove-Event
                continue
            }
        }

        $lines, $message = read_lines
        if (-not (activate_1c)) {continue}
        $wshell.SendKeys("{LEFT}"*20)
        $wshell.SendKeys("$(move_btw 1 $s.columns.code)")
        $lines | ForEach-Object {
            if ($_.remove -ne $null)
            {
                return
            }
            $wshell.SendKeys("{HOME}^f$([string]$_.code){ENTER}{ENTER}")
            $_.count = $_.count -replace '\.', '{RIGHT}'
            if ($_.in_pieces)
            {
                $wshell.SendKeys("$(move_btw $s.columns.code $s.columns.count){F2}$($_.count)+{F2}$(move_btw $s.columns.count $s.columns.code)")
            }
            else
            {
                $wshell.SendKeys("$(move_btw $s.columns.code $s.columns.cargo){F2}$($_.count)+{F2}$(move_btw $s.columns.cargo $s.columns.code)")
            }
            if ($_.warning -ne $null)
            {
                $wshell.SendKeys("$(move_btw $s.columns.code $s.columns.warning) $(move_btw $s.columns.warning $s.columns.code)")
            }
        }
	if ($lines.count -gt 0)
	{
	    $wshell.SendKeys("{ESC}")
            message "Набрано: $($lines.count)" "$message"
	}
	else
	{
	    message "Пустая заявка" "В выделенной области нет строк заявки"
	}
        #$Excel.Selection.Resize($Excel.Selection.Rows.Count, $Excel.Selection.Columns.Count+1).Select() > $null
        #$wshell.AppActivate((Get-Process Excel | Where-Object {$_.MainWindowHandle -eq $Excel.Hwnd}).Id) > $null

    Get-Event | Remove-Event
    }   
}
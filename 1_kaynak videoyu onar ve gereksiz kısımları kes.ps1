Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$global:fpickerBox = $null
$global:dpickerBox = $null

Function Get-FileName{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = "Kaynak dosyasý seçin"
    $OpenFileDialog.Filter = 'MP4 (*.mp4)|*.mp4|Herhangi dosya (*.*)|*.*'
    $OpenFileDialog.ShowDialog() | Out-Null
    return $OpenFileDialog.FileName
}

function iþlemip{
    $null = Read-Host 'Ýþlem iptal edildi... Çýkmak için herhangi bir tuþa basýn'
    exit
}

function askbeg{
    Param($form)

    $tvals = @('Baþlangýç:','Bitiþ:')
    $titles = @('Saat','Dakika','Saniye','MiliSn')
    $labels = @()
    $time = @()

    $form.Text = 'Kesim süreleri'
    $form.Size = New-Object System.Drawing.Size(380,245)
    $form.StartPosition = 'CenterScreen'

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(5,5)
    $label.Size = New-Object System.Drawing.Size(200,15)
    $label.Text = 'Kesim aralýklarýný girin:'
    $form.Controls.Add($label)

    foreach ($marker in $tvals){
        $labels += New-Object System.Windows.Forms.Label
        $mry = if ($marker -eq 'Bitiþ:') {80} else {30}
        $labels[$labels.length-1].Location = New-Object System.Drawing.Point(10,$mry)
        $labels[$labels.length-1].Size = New-Object System.Drawing.Size(60,15)
        $labels[$labels.length-1].Text = $marker
        $form.Controls.Add($labels[$labels.length-1])
        for ($i = 0; $i -lt 4; $i++){
            $labels += New-Object System.Windows.Forms.Label
            $labels[$labels.length-1].Location = New-Object System.Drawing.Point((20+$i*80),($mry+23))
            $labels[$labels.length-1].Size = New-Object System.Drawing.Size(40,15)
            $labels[$labels.length-1].Text = $titles[$i]
            $form.Controls.Add($labels[$labels.length-1])

            $time += New-Object System.Windows.Forms.NumericUpDown
            $time[$time.length-1].Location = New-Object System.Drawing.Point((60+$i*80),($mry+20))
            $time[$time.length-1].Size = New-Object System.Drawing.Size(40,15)
            $time[$time.length-1].Maximum = if ($i -eq 0) {23} elseif ($i -eq 3) {99} else {59}
            $form.Controls.Add($time[$time.length-1])
        }
    }
    #####

    $global:fpickerBox = New-Object System.Windows.Forms.TextBox
    $global:fpickerBox.Location = New-Object System.Drawing.Point(20,126)
    $global:fpickerBox.Size = New-Object System.Drawing.Size(250,23)
    $global:fpickerBox.Text = "kaynak.mp4"
    $form.Controls.Add($global:fpickerBox)

    $fpickerBut = New-Object System.Windows.Forms.Button
    $fpickerBut.Location = New-Object System.Drawing.Point(270,125)
    $fpickerBut.Size = New-Object System.Drawing.Size(80,22)
    $fpickerBut.Text = "Kaynak Seç"
    $form.Controls.Add($fpickerBut)

    $fpickerBut.Add_Click(
        {
            $temp = (Get-FileName)
            if($temp) {$global:fpickerBox.Text = $temp}
        }
    )

    #####

    $global:dpickerBox = New-Object System.Windows.Forms.TextBox
    $global:dpickerBox.Location = New-Object System.Drawing.Point(65,151)
    $global:dpickerBox.Size = New-Object System.Drawing.Size(250,23)
    $global:dpickerBox.Text = "raw_out.mp4"
    $global:dpickerBox.TextAlign = "center"
    $form.Controls.Add($global:dpickerBox)

    #####

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(140,175)
    $okButton.Size = New-Object System.Drawing.Size(100,23)
    $okButton.Text = 'Tamam'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    $form.Topmost = $true


    return $time
}

#############################################

$form = New-Object System.Windows.Forms.Form
$Form.FormBorderStyle = 'Fixed3D'
$num = askbeg($form)
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $x = $num | select -ExpandProperty Value
    $rawvid = $global:fpickerBox.Text
    $global:fpickerBox = $null
    $outname = $global:dpickerBox.Text
    $global:dpickerBox = $null
} else {
    iþlemip
}

if (-not (Test-Path -Path $rawvid -PathType Any)){
    Write-Host "Belirtilen dosya yok!"
    iþlemip
}

$icbas = New-TimeSpan -Hours $x[0] -Minutes $x[1] -Seconds ($x[2]+($x[3]/100))
$icbit = New-TimeSpan -Hours $x[4] -Minutes $x[5] -Seconds ($x[6]+($x[7]/100))

$sanbas = $icbas.Hours*60*60 + $icbas.Minutes*60 + $icbas.Seconds
$icsure = $icbit - $icbas

$icsure = $icsure.Hours*60*60 + $icsure.Minutes*60 + $icsure.Seconds

Write-Host "Kaynak video editlenmeye hazýrlanýyor"

$realt = (ffprobe -i `"$rawvid`" -show_entries format=duration -v quiet -of csv="p=0")

if (($sanbas -lt 0) -or ($icsure -lt 0) -or ($realt -lt $sanbas) -or ($realt -lt $icsure) ){
    Write-Host "Kesim süresi hatalý!"
    iþlemip
}

if (($sanbas -eq 0) -and ($icsure -eq 0)) {
    ffmpeg -loglevel quiet -stats -i `"$rawvid`" -c copy `"$outname`"
} elseif (($sanbas -ne 0) -and ($icsure -eq 0)) {
	ffmpeg -loglevel quiet -stats -ss ($sanbas -replace ',','.') -i `"$rawvid`" -c copy `"$outname`"
} elseif (($sanbas -eq 0) -and ($icsure -ne 0)){
	ffmpeg -loglevel quiet -stats -i `"$rawvid`" -c copy -t ($icsure -replace ',','.') `"$outname`"
} else {
    ffmpeg -loglevel quiet -stats -ss ($sanbas -replace ',','.') -i `"$rawvid`" -c copy -t ($icsure -replace ',','.') `"$outname`"
}

Pause
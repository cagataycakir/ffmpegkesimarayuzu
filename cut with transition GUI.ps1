Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Global:ffmpegData = @{}

function PSObjectToHash {
    [CmdletBinding()]
    [OutputType('hashtable')]
    param (
        [Parameter(ValueFromPipeline)]
        $InputObject
    )
    process
    {
        if ($null -eq $InputObject) { return $null }

        if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]){
            $collection = @(
                foreach ($object in $InputObject) { PSObjectToHash $object }
            )

            Write-Output -NoEnumerate $collection
        }elseif ($InputObject -is [psobject]){
            $hash = @{}

            foreach ($property in $InputObject.PSObject.Properties){
                $hash[$property.Name] = PSObjectToHash $property.Value
            }

            $hash
        }else{
            $InputObject
        }
    }
}

$getHash = {
    param ($file)
    Get-Content -Raw $file | ConvertFrom-Json | PSObjectToHash
}

function GUI ($w=650, $h=600){
    $form = New-Object System.Windows.Forms.Form
    $form.ShowIcon = $false
    $form.Text = 'Cut with transition'
    $form.MinimumSize = New-Object System.Drawing.Size($w,$h)
    $form.MaximumSize = New-Object System.Drawing.Size($w,[int]::MaxValue)
    $form.Size = New-Object System.Drawing.Size($w,$h)
    $form.StartPosition = 'CenterScreen'

    $form.Controls.Add((New-Object System.Windows.Forms.GroupBox)) ##[0] Trim menubar, labels and sections group
    $form.Controls.Add((New-Object System.Windows.Forms.Panel))    ##[1] Intro, outro and watermark options group
    $form.Controls.Add((New-Object System.Windows.Forms.Panel))    ##[2] Intro, content, watermark and outro endscreen file group
    $form.Controls.Add((New-Object System.Windows.Forms.Panel))    ##[3] Output name, format and enable group
    $form.Controls.Add((New-Object System.Windows.Forms.Panel))    ##[4] Video generation button group
    $form.Controls.Add((New-Object System.Windows.Forms.MenuStrip))##[5] Options menubar group

    #function that pops up file selection dialog box
    $getFileName = {
        param($filter)
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        #$OpenFileDialog.Title = "Select file"
        $OpenFileDialog.Filter = $filter
        $OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.FileName
    }.GetNewClosure()

    ##these functions needed for file selector type for outro card/video
    $ofc = {
        $t = Invoke-Command $getFileName -ArgumentList 'PNG (*.png)|*.png|Any file (*.*)|*.*'
        if($t) {$form.Controls[2].Controls[2].Controls[0].Text = $t}
    }.GetNewClosure()
    $ofv = {
        $t = Invoke-Command $getFileName -ArgumentList 'MP4 (*.mp4)|*.mp4|Any file (*.*)|*.*'
        if($t) {$form.Controls[2].Controls[2].Controls[0].Text = $t}
    }.GetNewClosure()

    #fills options from json data
    $fillOpt = {
        param($hash)

        #intro data
        if ($hash["Intro"]){        
            $form.Controls[1].Controls[2].Controls[0].Value = $hash["Intro"]["Intro Transition"]
            $form.Controls[2].Controls[0].Controls[0].Text = $hash["Intro"]["Intro File"]
            $form.Controls[3].Controls[2].Checked = $hash["Intro"]["Enabled"]
        }
        #source data
        if ($hash["Source File"]){
            $form.Controls[2].Controls[1].Controls[0].Text = $hash["Source File"]
        }
        #outro data
        if ($hash["Outro Card"]){
            $form.Controls[1].Controls[0].Controls[-1].SelectedIndex = 0
            $form.Controls[1].Controls[0].Controls[0].Controls[2].Controls[0].Value = $hash["Outro Card"]["Outro Transition"]
            $form.Controls[1].Controls[0].Controls[0].Controls[1].Controls[0].Value = $hash["Outro Card"]["Outro Length"]
            $form.Controls[1].Controls[0].Controls[0].Controls[0].Controls[2].Value = $hash["Outro Card"]["PosX"]
            $form.Controls[1].Controls[0].Controls[0].Controls[0].Controls[0].Value = $hash["Outro Card"]["PosY"]
            $form.Controls[1].Controls[0].Controls[0].Controls[0].Controls[5].Value = $hash["Outro Card"]["Size"]
            $form.Controls[2].Controls[2].Controls[0].Text = $hash["Outro Card"]["Outro File"]
            $form.Controls[3].Controls[1].Checked = $hash["Outro Card"]["Enabled"]
        }
        if ($hash["Outro Video"]){
            $form.Controls[1].Controls[0].Controls[-1].SelectedIndex = 1
            $form.Controls[2].Controls[2].Controls[0].Text = $hash["Outro Video"]["Outro File"]
            $form.Controls[1].Controls[0].Controls[0].Controls[0].Controls[0].Value = $hash["Outro Video"]["Outro Transition"]
            $form.Controls[3].Controls[1].Checked = $hash["Outro Video"]["Enabled"]
        }
        #watermark data
        if ($hash["Watermark Corner"]){
            $form.Controls[1].Controls[1].Controls[-1].SelectedIndex = 0
            $form.Controls[1].Controls[1].Controls[1].Controls[0].Value = $hash["Watermark Corner"]["Opacity"]
            $form.Controls[1].Controls[1].Controls[1].Controls[3].Value = $hash["Watermark Corner"]["Size"]
            $form.Controls[1].Controls[1].Controls[0].Controls[0].SelectedIndex = $hash["Watermark Corner"]["Corner"]
            $form.Controls[2].Controls[3].Controls[0].Text = $hash["Watermark Corner"]["Watermark File"]
            $form.Controls[3].Controls[0].Checked = $hash["Watermark Corner"]["Enabled"]
        }
        if ($hash["Watermark Position"]){
            $form.Controls[1].Controls[1].Controls[-1].SelectedIndex = 1
            $form.Controls[1].Controls[1].Controls[1].Controls[0].Value = $hash["Watermark Position"]["Opacity"]
            $form.Controls[1].Controls[1].Controls[1].Controls[3].Value = $hash["Watermark Position"]["Size"]
            $form.Controls[1].Controls[1].Controls[0].Controls[2].Value = $hash["Watermark Position"]["PosX"]
            $form.Controls[1].Controls[1].Controls[0].Controls[0].Value = $hash["Watermark Position"]["PosY"]
            $form.Controls[2].Controls[3].Controls[0].Text = $hash["Watermark Position"]["Watermark File"]
            $form.Controls[3].Controls[0].Checked = $hash["Watermark Position"]["Enabled"]
        }
        #output data
        if ($hash["Output"]){
            $form.Controls[3].Controls[3].SelectedIndex = $hash["Output"]["Output Ratio"]
            $form.Controls[3].Controls[4].Text = $hash["Output"]["Output File"]
            $form.Controls[3].Controls[5].SelectedItem = $hash["Output"]["Output Format"]
        }

    }.GetNewClosure()
    #gathers options into a hashtable object
    $gatherOpt = {
        $json = @{}
        #intro data
        $json += @{"Intro" = @{
            "Intro Transition" = $form.Controls[1].Controls[2].Controls[0].Value;
            "Intro File" = $form.Controls[2].Controls[0].Controls[0].Text;
            "Enabled" = $form.Controls[3].Controls[2].Checked
            }}
        #source data
        $json += @{ "Source File" = $form.Controls[2].Controls[1].Controls[0].Text }
        #outro data
        Switch($form.Controls[1].Controls[0].Controls[-1].SelectedIndex){
            0{$json += @{"Outro Card" = @{
                "Outro Transition" = $form.Controls[1].Controls[0].Controls[0].Controls[2].Controls[0].Value;
                "Outro Length" = $form.Controls[1].Controls[0].Controls[0].Controls[1].Controls[0].Value;
                "PosX" = $form.Controls[1].Controls[0].Controls[0].Controls[0].Controls[2].Value;
                "PosY" = $form.Controls[1].Controls[0].Controls[0].Controls[0].Controls[0].Value;
                "Size" = $form.Controls[1].Controls[0].Controls[0].Controls[0].Controls[5].Value;
                "Outro File" = $form.Controls[2].Controls[2].Controls[0].Text;
                "Enabled" = $form.Controls[3].Controls[1].Checked
                }}
            }
            1{$json += @{"Outro Video" = @{
                "Outro Transition" = $form.Controls[1].Controls[0].Controls[0].Controls[0].Controls[0].Value;
                "Outro File" = $form.Controls[2].Controls[2].Controls[0].Text;
                "Enabled" = $form.Controls[3].Controls[1].Checked
                }}
            }
        }
        #watermark data
        Switch($form.Controls[1].Controls[1].Controls[-1].SelectedIndex){
            0{$json += @{"Watermark Corner" = @{
                "Opacity" = $form.Controls[1].Controls[1].Controls[1].Controls[0].Value;
                "Size" = $form.Controls[1].Controls[1].Controls[1].Controls[3].Value;
                "Corner" = $form.Controls[1].Controls[1].Controls[0].Controls[0].SelectedIndex;
                "Watermark File" = $form.Controls[2].Controls[3].Controls[0].Text;
                "Enabled" = $form.Controls[3].Controls[0].Checked
                }}
            }
            1{$json += @{"Watermark Position" = @{
                "Opacity" = $form.Controls[1].Controls[1].Controls[1].Controls[0].Value;
                "Size" = $form.Controls[1].Controls[1].Controls[1].Controls[3].Value;
                "PosX" = $form.Controls[1].Controls[1].Controls[0].Controls[2].Value;
                "PosY" = $form.Controls[1].Controls[1].Controls[0].Controls[0].Value;
                "Watermark File" = $form.Controls[2].Controls[3].Controls[0].Text;
                "Enabled" = $form.Controls[3].Controls[0].Checked
                }}
            }
        }
        #output data
        $json += @{"Output" = @{
            "Output Ratio" = $form.Controls[3].Controls[3].SelectedIndex;
            "Output File" = $form.Controls[3].Controls[4].Text;
            "Output Format" = $form.Controls[3].Controls[5].SelectedItem
            }}

        $json
    }.GetNewClosure()

    #inserts a trim point
    $addTrim = {
            $trimList = $form.Controls[0].Controls[0]
            $bar = New-Object System.Windows.Forms.Panel -Property @{
                        Dock=[System.Windows.Forms.DockStyle]::Top;
                        Width=$trimList.Width-$trimList.Margin.All*2;
                        Height=20;
                    }
            $trimList.Controls.Add($bar)
            if($trimList.Controls.Count -gt 1){
                $trimList.Controls[-2].Controls[1].Controls[-1].Enabled=$true
            }

            $delbut = New-Object System.Windows.Forms.Button -Property @{
                        Dock=[System.Windows.Forms.DockStyle]::Left;
                        Width=20;
                        Text="-"
                    }
            
            $bar.Controls.Add($delbut)

            $numHolder = New-Object System.Windows.Forms.Panel -Property @{
                        Dock=[System.Windows.Forms.DockStyle]::Right;
                        Width=($bar.Width-[System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth-6)*14/15;
                    }
            $bar.Controls.Add($numHolder)

            $bar.Controls.Add((New-Object System.Windows.Forms.Panel -Property @{
                        Dock=[System.Windows.Forms.DockStyle]::Right;
                        Width=[System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth
                    }))

            function temp {
                foreach($lim in @(23,59,59,0.99)){
                    $numHolder.Controls.Add((New-Object System.Windows.Forms.NumericUpDown -Property @{
                        Width=$numHolder.Width/10;
                        Dock=[System.Windows.Forms.DockStyle]::Right;
                        TextAlign="Center";
                        Maximum=$lim;
                        DecimalPlaces=if($lim -eq 0.99){2}else{0}
                        Increment=if($lim -eq 0.99){0.01}else{1}
                    }))
                }
            }

            temp
            $numHolder.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
                        Width=$numHolder.Width/10;
                        Dock=[System.Windows.Forms.DockStyle]::Right;
                        Text="-";
                        TextAlign=[System.Drawing.ContentAlignment]::MiddleCenter
                    }))
            temp
            $numHolder.Controls.Add((New-Object System.Windows.Forms.Panel -Property @{Width=5;Dock=[System.Windows.Forms.DockStyle]::Right}))
            $numHolder.Controls.Add((New-Object System.Windows.Forms.NumericUpDown -Property @{
                        Width=$numHolder.Width/10-5;
                        Dock=[System.Windows.Forms.DockStyle]::Right;
                        TextAlign="Center";
                        Maximum=10;
                        DecimalPlaces=2;
                        Increment=0.02;
                        Enabled=$false
                    }))

            #init button to remove holder frame
            $thisFade = $numHolder.Controls[-1]
            $delbut.Add_Click({
                if(-not $thisFade.Enabled -and $trimList.Controls.Count -gt 1){$trimList.Controls[-2].Controls[1].Controls[-1].Enabled=$false}
                $trimList.Controls.Remove($this.Parent)
            }.GetNewClosure())
        }.GetNewClosure()

    function init0 {
        $form.Controls[0].Dock = [System.Windows.Forms.DockStyle]::Fill
        $form.Controls[0].Text = "Source video trim points"
        $form.Controls[0].Controls.Add((New-Object System.Windows.Forms.FlowLayoutPanel -Property @{
            Dock=[System.Windows.Forms.DockStyle]::Fill;
            WrapContents = $true;
        }))

        $form.Controls[0].Controls[0].HorizontalScroll.Maximum = 0;
        $form.Controls[0].Controls[0].AutoScroll = $true;

        ##Init labels
        $form.Controls[0].Controls.Add((New-Object System.Windows.Forms.Panel -Property @{Height=40;Dock=[System.Windows.Forms.DockStyle]::Top}))
        $labelhol = New-Object System.Windows.Forms.Panel
        $form.Controls[0].Controls[1].Controls.Add($labelhol)
        $form.Controls[0].Controls[1].Controls.Add((New-Object System.Windows.Forms.Panel -Property @{
                        Dock=[System.Windows.Forms.DockStyle]::Right;
                        Width=[System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth;
                    }))

        $labelhol.Dock = [System.Windows.Forms.DockStyle]::Right
        $labelhol.Width = ($labelhol.Parent.Width-[System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth)*14/15

        foreach($i in ("Start","End"," ")){
            $labelhol.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
                Dock=[System.Windows.Forms.DockStyle]::Right;
                TextAlign=$(if($i -eq "Fade"){[System.Drawing.ContentAlignment]::BottomCenter}else{[System.Drawing.ContentAlignment]::MiddleCenter});
                Text=$i;
                Width=($labelhol.Width/11*$(if($i -eq " "){1}else{5}))
            }))
        }
        $divider = New-Object System.Windows.Forms.Panel -Property @{Dock=[System.Windows.Forms.DockStyle]::Bottom; Height=20}
        foreach($_ in ($true,$false)){
            foreach($i in ("Hour","Min","Sec","MilSec")){
                $divider.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
                    Dock=[System.Windows.Forms.DockStyle]::Right;
                    TextAlign=[System.Drawing.ContentAlignment]::MiddleCenter;
                    Text=$i;
                    Width=($labelhol.Width/10)
                }))
           }
           $divider.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
             Dock=[System.Windows.Forms.DockStyle]::Right;
             TextAlign=$(if($_){[System.Drawing.ContentAlignment]::MiddleCenter}else{[System.Drawing.ContentAlignment]::BottomCenter});
             Text=$(if($_){"-"}else{"Fade"});
                Width=($labelhol.Width/10)
            }))
        }
        $labelhol.Controls.Add($divider)

        ##Init Menubar And Functions
        $getFileName = $getFileName
        $addTrim = $addTrim
        $getHash = $getHash
        $trimList = $form.Controls[0].Controls[0].Controls
        $form.Controls[0].Controls.Add((New-Object System.Windows.Forms.MenuStrip))
        $null = $form.Controls[0].Controls[2].Items.Add(( New-Object System.Windows.Forms.ToolStripMenuItem -Property @{ Text="Import trim points" } ))
        $form.Controls[0].Controls[2].Items[-1].Add_Click({
            $filename = Invoke-Command $getFileName -ArgumentList 'Json (*.json)|*.json'
            if($filename) {
                $json = (Invoke-Command -ScriptBlock $getHash -ArgumentList $filename)["data"]
                $trimList.Clear()

                foreach ($i in $json){
                    Invoke-Command $addTrim
                    $trimList[-1].Controls[1].Controls[0].Value = $i["Start"]["Hour"]
                    $trimList[-1].Controls[1].Controls[1].Value = $i["Start"]["Min"]
                    $trimList[-1].Controls[1].Controls[2].Value = $i["Start"]["Sec"]
                    $trimList[-1].Controls[1].Controls[3].Value = $i["Start"]["MiliSec"]
                    $trimList[-1].Controls[1].Controls[5].Value = $i["End"]["Hour"]
                    $trimList[-1].Controls[1].Controls[6].Value = $i["End"]["Min"]
                    $trimList[-1].Controls[1].Controls[7].Value = $i["End"]["Sec"]
                    $trimList[-1].Controls[1].Controls[8].Value = $i["End"]["MiliSec"]
                    $trimList[-1].Controls[1].Controls[10].Value = $i["Fade"]

                }
            }
        }.GetNewClosure())
        $null = $form.Controls[0].Controls[2].Items.Add(( New-Object System.Windows.Forms.ToolStripMenuItem -Property @{ Text="Save trim points" } ))
        $form.Controls[0].Controls[2].Items[-1].Add_Click({
            $trimList = $form.Controls[0].Controls[0].Controls

            $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
            $SaveFileDialog.FileName = "Trim points"
            $SaveFileDialog.Filter = 'Json (*.json)|*.json'
            $SaveFileDialog.ShowDialog() #| Out-Null
            if($SaveFileDialog.FileName){
                function gatherSec {
                    $out = @()
                    foreach($i in $trimList){
                        $out += @{
                            "Start"=@{
                                "Hour"=$i.Controls[1].Controls[0].Value;
                                "Min"=$i.Controls[1].Controls[1].Value;
                                "Sec"=$i.Controls[1].Controls[2].Value;
                                "MiliSec"=$i.Controls[1].Controls[3].Value
                            };
                            "End"=@{
                                "Hour"=$i.Controls[1].Controls[5].Value;
                                "Min"=$i.Controls[1].Controls[6].Value;
                                "Sec"=$i.Controls[1].Controls[7].Value;
                                "MiliSec"=$i.Controls[1].Controls[8].Value
                            };
                            "Fade"=$i.Controls[1].Controls[10].Value
                        }
                    }
                    $out
                }
                $json = @{"data"=gatherSec}
                ConvertTo-Json $json -Depth 3 | Out-File $SaveFileDialog.FileName
            }
        })
        $null = $form.Controls[0].Controls[2].Items.Add(( New-Object System.Windows.Forms.ToolStripMenuItem -Property @{ Text="Add trim point"; Alignment="Right" } ))
        $form.Controls[0].Controls[2].Items[-1].Add_Click($addTrim)
    }
    function init1 {
        $form.Controls[1].Dock = [System.Windows.Forms.DockStyle]::Top
        $div = ($w-12)/3

        function initOut {
            $outroBox = New-Object System.Windows.Forms.GroupBox -Property @{Dock=[System.Windows.Forms.DockStyle]::Right; Text="Outro Options"; Width=$div}
            $changingOpt = New-Object System.Windows.Forms.Panel -Property @{Dock=[System.Windows.Forms.DockStyle]::Fill}

            $ooc = {
                $changingOpt.Controls.Add((New-Object System.Windows.Forms.Panel -Property @{Dock=[System.Windows.Forms.DockStyle]::Top; Height=20}))
                $changingOpt.Controls.Add((New-Object System.Windows.Forms.Panel -Property @{Dock=[System.Windows.Forms.DockStyle]::Top; Height=20}))
                $changingOpt.Controls.Add((New-Object System.Windows.Forms.Panel -Property @{Dock=[System.Windows.Forms.DockStyle]::Top; Height=20}))

                $changingOpt.Controls[2].Controls.Add((New-Object System.Windows.Forms.NumericUpDown -Property @{Dock=[System.Windows.Forms.DockStyle]::Right; AutoSize=$true; Maximum=10}))
                $changingOpt.Controls[2].Controls.Add((New-Object System.Windows.Forms.Label -Property @{
                    Dock=[System.Windows.Forms.DockStyle]::Left;
                    Text="Outro Transition"
                    TextAlign=[System.Drawing.ContentAlignment]::MiddleLeft;
                    AutoSize=$true
                }))

                $changingOpt.Controls[1].Controls.Add((New-Object System.Windows.Forms.NumericUpDown -Property @{Dock=[System.Windows.Forms.DockStyle]::Right; Minimum=1; Value=15; AutoSize=$true}))
                $changingOpt.Controls[1].Controls.Add((New-Object System.Windows.Forms.Label -Property @{
                    Dock=[System.Windows.Forms.DockStyle]::Left;
                    Text="Outro Length"
                    TextAlign=[System.Drawing.ContentAlignment]::MiddleLeft;
                    AutoSize=$true
                }))

                foreach ($i in ("Y","X")){
                    $changingOpt.Controls[0].Controls.Add((New-Object System.Windows.Forms.NumericUpDown -Property @{Dock=[System.Windows.Forms.DockStyle]::Left; AutoSize=$true; Maximum=1920}))
                    $changingOpt.Controls[0].Controls.Add((New-Object System.Windows.Forms.Label -Property @{
                        Dock=[System.Windows.Forms.DockStyle]::Left;
                        Text=$i;
                        TextAlign=[System.Drawing.ContentAlignment]::MiddleLeft;
                        AutoSize=$true
                    }))
                }
                $changingOpt.Controls[0].Controls.Add((New-Object System.Windows.Forms.Label -Property @{
                    Dock=[System.Windows.Forms.DockStyle]::Right;
                    Text="Size";
                    TextAlign=[System.Drawing.ContentAlignment]::MiddleLeft;
                    AutoSize=$true
                }))
                $changingOpt.Controls[0].Controls.Add((New-Object System.Windows.Forms.NumericUpDown -Property @{Dock=[System.Windows.Forms.DockStyle]::Right; AutoSize=$true; Value=25; Minimum=1; DecimalPlaces=1}))

            }.GetNewClosure()
            $oov = {
                $changingOpt.Controls.Add((New-Object System.Windows.Forms.Panel -Property @{Dock=[System.Windows.Forms.DockStyle]::Top; Height=20}))
                $changingOpt.Controls[0].Controls.Add((New-Object System.Windows.Forms.NumericUpDown -Property @{Dock=[System.Windows.Forms.DockStyle]::Right; AutoSize=$true}))
                $changingOpt.Controls[0].Controls.Add((New-Object System.Windows.Forms.Label -Property @{
                    Dock=[System.Windows.Forms.DockStyle]::Left;
                    Text="Outro Transition"
                    TextAlign=[System.Drawing.ContentAlignment]::MiddleLeft;
                    AutoSize=$true
                }))
            }.GetNewClosure()
            Invoke-Command $ooc

            $optController = New-Object System.Windows.Forms.ComboBox
            $optController.Dock = [System.Windows.Forms.DockStyle]::Top
            $optController.Items.AddRange(@("Outro Card","Outro Video"))
            $optController.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
            $optController.SelectedIndex = 0

            $frmptr = $form.Controls[2]

            #this needed to point to closures
            $ofc=$ofc
            $ofv=$ofv

            $optController.Add_SelectedIndexChanged({
                $frmptr = $frmptr.Controls[2]
                $frmptr.Text = $optController.SelectedItem
                $frmptr.Controls[0].Text = ""
                
                $changingOpt.Controls.Clear()
                Switch($optController.SelectedIndex){
                    0 {
                        Invoke-Command $ooc
                        $frmptr.Controls[1].Remove_Click($ofv)
                        $frmptr.Controls[1].Add_Click($ofc)
                    }
                    1 {
                        Invoke-Command $oov
                        $frmptr.Controls[1].Remove_Click($ofc)
                        $frmptr.Controls[1].Add_Click($ofv)
                    }
                }
            }.GetNewClosure())

            $outroBox.Controls.Add($changingOpt)
            $outroBox.Controls.Add($optController)
            $form.Controls[1].Controls.Add($outroBox)
        }
        function initWater{
            $wmarkBox = New-Object System.Windows.Forms.GroupBox -Property @{Dock=[System.Windows.Forms.DockStyle]::Right; Text="Watermark Options"; Width=$div}
            $sameOpt = New-Object System.Windows.Forms.Panel -Property @{Dock=[System.Windows.Forms.DockStyle]::Top; Height=20}
            $changingOpt = New-Object System.Windows.Forms.Panel -Property @{Dock=[System.Windows.Forms.DockStyle]::Fill}

            $sameOpt.Controls.Add((New-Object System.Windows.Forms.NumericUpDown -Property @{Dock=[System.Windows.Forms.DockStyle]::Left; AutoSize=$true; Value=100; Minimum=1}))
            $sameOpt.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
                Dock=[System.Windows.Forms.DockStyle]::Left;
                Text="Opacity";
                TextAlign=[System.Drawing.ContentAlignment]::MiddleLeft;
                AutoSize=$true
            }))
            $sameOpt.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
                Dock=[System.Windows.Forms.DockStyle]::Right;
                Text="Size"
                TextAlign=[System.Drawing.ContentAlignment]::MiddleLeft;
                AutoSize=$true
            }))
            $sameOpt.Controls.Add((New-Object System.Windows.Forms.NumericUpDown -Property @{Dock=[System.Windows.Forms.DockStyle]::Right; AutoSize=$true; Value=25; Minimum=1}))

            $wmc = {
                $changingOpt.Controls.Add((New-Object System.Windows.Forms.ComboBox -Property @{Dock=[System.Windows.Forms.DockStyle]::Top; Height=20}))

                $changingOpt.Controls[0].Items.AddRange(@("Top Left","Top Right","Bottom Left","Bottom Right"))
                $changingOpt.Controls[0].DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
                $changingOpt.Controls[0].SelectedIndex = 3

            }.GetNewClosure()
            Invoke-Command $wmc

            $optController = New-Object System.Windows.Forms.ComboBox
            $optController.Dock = [System.Windows.Forms.DockStyle]::Top
            $optController.Items.AddRange(@("Corner","Position"))
            $optController.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
            $optController.SelectedIndex = 0

            $optController.Add_SelectedIndexChanged({
                $wmp = {
                    foreach($i in ("Y", "X")){
                        $changingOpt.Controls.Add((New-Object System.Windows.Forms.NumericUpDown -Property @{Dock=[System.Windows.Forms.DockStyle]::Left; AutoSize=$true}))
                        $changingOpt.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
                            Dock=[System.Windows.Forms.DockStyle]::Left;
                            Text=$i
                            TextAlign=[System.Drawing.ContentAlignment]::MiddleLeft;
                            AutoSize=$true
                        }))
                    }
                }
                
                $changingOpt.Controls.Clear()
                Switch($optController.SelectedIndex){
                    0 { Invoke-Command $wmc }
                    1 { Invoke-Command $wmp }
                }
            }.GetNewClosure())

            $wmarkBox.Controls.Add($changingOpt)
            $wmarkBox.Controls.Add($sameOpt)
            $wmarkBox.Controls.Add($optController)
            $form.Controls[1].Controls.Add($wmarkBox)
        }
        function initIntro{
            $introBox = New-Object System.Windows.Forms.GroupBox -Property @{Dock=[System.Windows.Forms.DockStyle]::Left; Text="Intro Options"; Width=$div}
            $introBox.Controls.Add((New-Object System.Windows.Forms.NumericUpDown -Property @{Dock=[System.Windows.Forms.DockStyle]::Right; AutoSize=$true; Maximum=10}))
            $introBox.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
                Dock=[System.Windows.Forms.DockStyle]::Left;
                Text="Intro Transition"
                TextAlign=[System.Drawing.ContentAlignment]::MiddleLeft;
                AutoSize=$true
            }))
            $form.Controls[1].Controls.Add($introBox)
        }

        initOut
        initWater
        initIntro
    }
    function init2 ($hp=50){
        $form.Controls[2].Dock = [System.Windows.Forms.DockStyle]::Top
        $form.Controls[2].Height = $hp
        $wp = $form.Controls[2].Width

        $getFileName = $getFileName

        foreach ($l in ("Intro Video","Source Video","Outro Card","Watermark")){
            $temp = New-Object System.Windows.Forms.GroupBox -Property @{Dock=[System.Windows.Forms.DockStyle]::Right; Width=($wp/4);Text=$l}
            $temp.Controls.Add((New-Object System.Windows.Forms.TextBox -Property @{
                Location=(New-Object System.Drawing.Point(0,($hp/3)));
                Width=($wp*2/4/3);
                Height=($hp/3)}
            ))
            $temp.Controls.Add((New-Object System.Windows.Forms.Button -Property @{
                Location=(New-Object System.Drawing.Point(($wp*2/4/3),($hp/3)));
                Width=($wp/4/3);
                Height=($hp/3);
                Text="File"}
            ))

            switch ($l){
                "Intro Video" { $temp.Controls[1].Add_Click( {
                    $t = Invoke-Command $getFileName -ArgumentList 'MP4 (*.mp4)|*.mp4|Any file (*.*)|*.*'
                    if($t) {$temp.Controls[0].Text = $t}
                }.GetNewClosure() )}
                "Source Video" { $temp.Controls[1].Add_Click( {
                    $t = Invoke-Command $getFileName -ArgumentList 'MP4 (*.mp4)|*.mp4|Any file (*.*)|*.*'
                    if($t) {$temp.Controls[0].Text = $t}
                }.GetNewClosure() )}
                "Outro Card" { $temp.Controls[1].Add_Click($ofc) }
                "Watermark" { $temp.Controls[1].Add_Click( {
                    $t = Invoke-Command $getFileName -ArgumentList 'PNG (*.png)|*.png|Any file (*.*)|*.*|AVI (*.avi)|*.avi'
                    if($t) {$temp.Controls[0].Text = $t}
                }.GetNewClosure() )}
            }

            $form.Controls[2].Controls.Add($temp)
        }
    }
    function init3 {
        $temp = New-Object System.Windows.Forms.ComboBox -Property @{
            Dock = [System.Windows.Forms.DockStyle]::Left;
            DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
            FlatStyle = [System.Windows.Forms.FlatStyle]::Standard;
            AutoSize=$true;
            Width=80;
        }

        foreach ($i in ("Add Watermark","Add Outro","Add Intro")){
            $temp = New-Object System.Windows.Forms.CheckBox -Property @{
                        Dock=[System.Windows.Forms.DockStyle]::Left;
                        Text=$i;
                        Checked=$true;
                        AutoSize=$true;
                    }
            Switch($i){
                "Add Watermark" {
                    $formfile = $form.Controls[2].Controls[3]
                    $formopt = $form.Controls[1].Controls[1]
                }
                "Add Outro" {
                    $formfile = $form.Controls[2].Controls[2]
                    $formopt = $form.Controls[1].Controls[0]
                }
                "Add Intro" {
                    $formfile = $form.Controls[2].Controls[0]
                    $formopt = $form.Controls[1].Controls[2]
                }
            }
            $temp.Add_CheckStateChanged({
                $formopt.Enabled = $formfile.Enabled = if($temp.Checked) {$true} else {$false}
            }.GetNewClosure())
            $form.Controls[3].Controls.Add($temp)
        }

        $temp = New-Object System.Windows.Forms.ComboBox -Property @{
            Dock = [System.Windows.Forms.DockStyle]::Right;
            DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
            FlatStyle = [System.Windows.Forms.FlatStyle]::Standard;
            AutoSize=$true;
            Width=100;
        }
        $temp.Items.AddRange(@("Same as source","1080p 30fps","1080p 60fps"))
        $temp.SelectedIndex = 1

        $form.Controls[3].Controls.Add($temp)

        $form.Controls[3].Controls.Add((
            New-Object System.Windows.Forms.TextBox -Property @{
                Dock=[System.Windows.Forms.DockStyle]::Right;
                Text="Output";
                TextAlign="Right"
            }
        ))

        $form.Controls[3].Dock = [System.Windows.Forms.DockStyle]::Bottom
        $form.Controls[3].Height = 20

        $temp = New-Object System.Windows.Forms.ComboBox -Property @{
            Dock = [System.Windows.Forms.DockStyle]::Right;
            DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
            FlatStyle = [System.Windows.Forms.FlatStyle]::Standard;
            AutoSize=$true;
            Width=60;
        }
        $temp.Items.AddRange(@(".mp4",".mov",".avi"))
        $temp.SelectedIndex = 0

        $form.Controls[3].Controls.Add($temp)
    }
    function init4 ($hp=40){
        $form.Controls[4].Dock = [System.Windows.Forms.DockStyle]::Bottom
        $form.Controls[4].Height = $hp*3/2
        $form.Controls[4].Controls.Add((
            New-Object System.Windows.Forms.Button -Property @{
                Dock=[System.Windows.Forms.DockStyle]::Bottom;
                Height = $hp;
                Text="Generate Video";
                DialogResult = [System.Windows.Forms.DialogResult]::OK
                }
        ))
        $gatherOpt = $gatherOpt
        $trimptr = $form.Controls[0].Controls[0].Controls
        $form.Controls[4].Controls[-1].Add_Click({
            if ( $trimptr.Count -gt 0){
                $trim = @()
                foreach($i in $trimptr){
                    $trim += @{
                        "Start"=@{
                            "Hour"=$i.Controls[1].Controls[0].Value;
                            "Min"=$i.Controls[1].Controls[1].Value;
                            "Sec"=$i.Controls[1].Controls[2].Value;
                            "MiliSec"=$i.Controls[1].Controls[3].Value
                        };
                        "End"=@{
                            "Hour"=$i.Controls[1].Controls[5].Value;
                            "Min"=$i.Controls[1].Controls[6].Value;
                            "Sec"=$i.Controls[1].Controls[7].Value;
                            "MiliSec"=$i.Controls[1].Controls[8].Value
                        };
                        "Fade"=$i.Controls[1].Controls[10].Value
                    }
                }
                $Global:ffmpegData += @{ "TrimPoints" = $trim }
            }
            $Global:ffmpegData += @{ "Options" = Invoke-Command $gatherOpt }
        }.GetNewClosure())
    }
    function init5 {
        #fills options with gathered data
        $null = $form.Controls[5].Items.Add(( New-Object System.Windows.Forms.ToolStripMenuItem -Property @{ Text="Import template" }))
        $fillOpt = $fillOpt
        $getFileName = $getFileName
        $gatherOpt = $gatherOpt
        $getHash = $getHash
        $form.Controls[5].Items[-1].Add_Click({
            $myJson = Invoke-Command $getFileName -ArgumentList 'Json (*.json)|*.json'
            if($myJson) {
                $hash = Invoke-Command -ScriptBlock $getHash -ArgumentList $myJson
                Invoke-Command -ScriptBlock $fillOpt -ArgumentList $hash
            }
        }.GetNewClosure())

        #generates new json data from current settings
        $null = $form.Controls[5].Items.Add(( New-Object System.Windows.Forms.ToolStripMenuItem -Property @{ Text="Save template" }))
        $gatherOpt = $gatherOpt
        $form.Controls[5].Items[-1].Add_Click({
            $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
            $SaveFileDialog.FileName = "Template options"
            $SaveFileDialog.Filter = 'Json (*.json)|*.json'
            
            if($SaveFileDialog.ShowDialog() -eq 1){
                $json = Invoke-Command $gatherOpt
                ConvertTo-Json $json | Out-File $SaveFileDialog.FileName
            }
        }.GetNewClosure())
    }

    init0
    init1
    init2
    init3
    init4
    init5

    $form
}

function InvokeFFMPEG{
    ##checking if ffmpeg and ffprobe present
    $ff = if(get-command "ffmpeg" -errorAction SilentlyContinue){"ffmpeg"}elseif(Test-Path "ffmpeg\bin\ffmpeg.exe"){"ffmpeg\bin\ffmpeg.exe"}else{$null}
    $fp = if(get-command "ffprobe" -errorAction SilentlyContinue){"ffprobe"}elseif(Test-Path "ffmpeg\bin\ffprobe.exe"){"ffmpeg\bin\ffprobe.exe"}else{$null}
    if(-not $ff -or -not $fp){
        $null = Read-Host "Ffmpeg/ffprobe doesnt exist! Press enter to exit."
        exit
    }
    $ff += " -loglevel quiet -stats"
    $ff += " -y"

    ##linking source files
    $op = $Global:ffmpegData["Options"]
    if($op["Source File"] -and (Test-Path $op["Source File"])){
        $ff += " -i `"" + $op["Source File"] + "`"" #this video source will always be 0
    }else{
        $null = Read-Host "Source file doesn't exist or not given! Press enter to exit."
        exit
    }

    function overwriteConf{
        [System.Windows.Forms.MessageBox]::Show("Output file already exist. Do you want to overwrite?", "Overwrite warning",[System.Windows.Forms.MessageBoxButtons]::OKCancel) -eq "OK"
    }

    if((Test-Path  ($op["Output"]["Output File"] + $op["Output"]["Output Format"])) -and -not (overwriteConf)){
        $null = Read-Host "Overwrite has been cancelled. Press enter to exit."
        exit
    }

    $ctr = 1
    if($op["Intro"] -and $op["Intro"]["Intro File"] -and (Test-Path $op["Intro"]["Intro File"]) -and $op["Intro"]["Enabled"]){
        $ff += " -i `"" + $op["Intro"]["Intro File"]+ "`""
        $Global:ffmpegData += @{"fIntro" = $ctr++}
    }
    foreach($i in ( ("Outro", (" Video"," Card")), ("Watermark", (" Corner"," Position")) ) ){
        foreach($j in $i[1]){
            if($op[$i[0]+$j] -and $op[$i[0]+$j][$i[0]+" File"] -and (Test-Path $op[$i[0]+$j][$i[0]+" File"]) -and $op[$i[0]+$j]["Enabled"]){
                if($j -ne " Video") {$ff += " -loop 1"}
                $ff += " -i `"" + $op[$i[0]+$j][$i[0]+" File"]+ "`""
                $fmark = "f"+$i[0]
                $Global:ffmpegData += @{ $fmark = $ctr++ }
            }
        }
    }

    ##generating complex filter
    function ComplexFilter{
        $width = 1920
        $height = 1080
        
        #video format
        switch($op["Output"]["Output Ratio"]){
            0{
                $width,$height = Invoke-Expression ($fp + " -v error -show_entries stream=width,height -of default=noprint_wrappers=1 `""+ $op["Source File"]+ "`"")
                $width = [int] $width.split('=')[1]
                $height = [int] $height.split('=')[1]
                $fps = Invoke-Expression ($fp + " -v error -select_streams v -of default=noprint_wrappers=1:nokey=1 -show_entries stream=r_frame_rate `""+ $op["Source File"]+ "`"")
                #$fps = Invoke-Expression ($fps)
                $vidrat = "crop='if(gte(iw/"+$width+",ih/"+$height+"),ih/"+$height+"*"+$width+",iw)'"+":'if(gt(ih/"+$height+",iw/"+$width+"),iw/"+$width+"*"+$height+",ih)',scale="+$width+":"+$height+",fps="+$fps+",format=pix_fmts=yuva420p,setpts=PTS-STARTPTS,settb=AVTB"
            } #same as source
            1{$vidrat = "crop='if(gte(iw/16,ih/9),ih/9*16,iw)':'if(gt(ih/9,iw/16),iw/16*9,ih)',scale=1920:1080,fps=30,format=pix_fmts=yuva420p,setpts=PTS-STARTPTS,settb=AVTB"} #1080p 30 fps
            2{$vidrat = "crop='if(gte(iw/16,ih/9),ih/9*16,iw)':'if(gt(ih/9,iw/16),iw/16*9,ih)',scale=1920:1080,fps=60,format=pix_fmts=yuva420p,setpts=PTS-STARTPTS,settb=AVTB"} #1080p 60 fps
        }

        #set input video and audio streams
        $ret += "[0:v]"+$vidrat+"[v0];"+"[0:a]asettb=AVTB,loudnorm,loudnorm[a0];"
        
        if($Global:ffmpegData["fWatermark"]){
            $ival = $Global:ffmpegData["fWatermark"]
            $wmop = if($op["Watermark Corner"]){$op["Watermark Corner"]["Opacity"]}else{$op["Watermark Position"]["Opacity"]}
            $wmscal = if($op["Watermark Corner"]){$op["Watermark Corner"]["Size"]}else{$op["Watermark Position"]["Size"]}
            
            $ret += "["+$ival+":v]format=argb,colorchannelmixer=aa="+($wmop/100)+",scale=-1:"+($height * $wmscal/100)+"[v"+$ival+"];"
        }

        foreach($i in ("fIntro","fOutro")){
            if($Global:ffmpegData[$i]){
                $ival = $Global:ffmpegData[$i]
                $ret += "["+$ival+":v]"+$vidrat+"[v"+$ival+"];"

                if(-not ($i -eq "fOutro" -and $op["Outro Card"]) ){
                    $ret += "["+$ival+":a]asettb=AVTB,loudnorm,loudnorm[a"+$ival+"];"
                }
            }
        }

        #video cuts
        $sourceLen = [double] (Invoke-Expression ($fp + " -i `"" + $op["Source File"] + "`" -show_entries format=duration -v quiet -of csv=`"p=0 `""))
        $Global:cutLen = @()

        function trimPoint($data){
            $beg = $data["Start"]["Hour"]*60*60+$data["Start"]["Min"]*60+$data["Start"]["Sec"]+$data["Start"]["MiliSec"]
            $end = $data["End"]["Hour"]*60*60+$data["End"]["Min"]*60+$data["End"]["Sec"]+$data["End"]["MiliSec"]

            #errors
            if($beg -ne 0 -and $end -ne 0 -and $beg -ge $end){
                $null = Read-Host "Trimming error! Press enter to exit."
                exit
            }
            if($beg -ge $sourceLen -or $end -ge $sourceLen){
                $null = Read-Host "Trimming error! Press enter to exit."
                exit
            }

            if($beg -eq 0 -and $end -eq 0){
                $retv = "null"
                $reta = "anull"
                $Global:cutLen += ,$sourceLen
            }elseif($beg -eq 0 -and $end -ne 0){
                $retv ="trim=end="+$end+",setpts=PTS-STARTPTS"
                $reta ="atrim=end="+$end+",asetpts=PTS-STARTPTS"
                $Global:cutLen += ,$end
            }elseif($beg -ne 0 -and $end -eq 0){
                $retv = "trim=start="+$beg+",setpts=PTS-STARTPTS"
                $reta = "atrim=start="+$beg+",asetpts=PTS-STARTPTS"
                $Global:cutLen += ,($sourceLen - $beg)
            }elseif($beg -ne 0 -and $end -ne 0){
                $retv = "trim=start="+$beg+":end="+$end+",setpts=PTS-STARTPTS"
                $reta = "atrim=start="+$beg+":end="+$end+",asetpts=PTS-STARTPTS"
                $Global:cutLen += ,($end - $beg)
            }
            $retv, $reta
        }

        #   curlen will be used at the other parts of the main function
        if($Global:ffmpegData["TrimPoints"].Count -eq 0){
            $ret += "[v0]null[vCut];[a0]anull[aCut];"
            $curlen = $sourceLen
        }elseif($Global:ffmpegData["TrimPoints"].Count -eq 1){
            $trv,$tra = trimPoint $Global:ffmpegData["TrimPoints"][0]
            $ret += "[v0]"+$trv+"[vCut];[a0]"+$tra+"[aCut];"
            $curlen = $Global:cutLen[0]
        }else{
            #splits
            $vtem="[v0]split="+$Global:ffmpegData["TrimPoints"].Count
            $atem="[a0]asplit="+$Global:ffmpegData["TrimPoints"].Count

            for($i=0; $i -lt $Global:ffmpegData["TrimPoints"].Count; $i++){
                $vtem += "[v0_"+$i+"]"
                $atem += "[a0_"+$i+"]"
            }
            $vtem += ";"
            $atem += ";"

            $ret += $vtem + $atem

            #cuts
            for($i=0; $i -lt $Global:ffmpegData["TrimPoints"].Count; $i++){
                $trv,$tra = trimPoint $Global:ffmpegData["TrimPoints"][$i]
                $ret += "[v0_"+$i+"]"+$trv+"[v0_"+$i+"];"
                $ret += "[a0_"+$i+"]"+$tra+"[a0_"+$i+"];"
            }

            #fades
            $curlen = $Global:cutLen[0]

            $fdur = $Global:ffmpegData["TrimPoints"][0]["Fade"]
            if($fdur -gt 0){
                $ret += "[v0_0][v0_1]xfade=transition=fade:duration="+$fdur+":offset="+($curlen-$fdur)+"[v0_0_1];"
                $ret += "[a0_0][a0_1]acrossfade=d="+$fdur+"[a0_0_1];"
                $curlen += $Global:cutLen[1] - $fdur
            }else{
                $ret += "[v0_0][v0_1]concat[v0_0_1];"
                $ret += "[a0_0][a0_1]concat=v=0:a=1[a0_0_1];"
                $curlen += $Global:cutLen[1]
            }

            for($i=1; $i -lt $Global:ffmpegData["TrimPoints"].Count-1; $i++){
                $fdur = $Global:ffmpegData["TrimPoints"][$i]["Fade"]

                if($fdur -gt 0){
                    $ret += "[v0_0_"+$i+"]"+"[v0_"+($i+1)+"]"+"xfade=transition=fade:duration="+$fdur+":offset="+($curlen-$fdur)+"[v0_0_"+($i+1)+"];"
                    $ret += "[a0_0_"+$i+"]"+"[a0_"+($i+1)+"]"+"acrossfade=d="+$fdur+"[a0_0_"+($i+1)+"];"
                    $curlen += $Global:cutLen[$i+1] - $fdur
                }else{
                    $ret += "[v0_0_"+$i+"]"+"[v0_"+($i+1)+"]concat[v0_0_"+($i+1)+"];"
                    $ret += "[a0_0_"+$i+"]"+"[a0_"+($i+1)+"]concat=v=0:a=1[a0_0_"+($i+1)+"];"
                    $curlen += $Global:cutLen[$i+1]
                }
            }

            $ret += "[v0_0_"+ ($Global:ffmpegData["TrimPoints"].Count-1) +"]null[vCut];"
            $ret += "[a0_0_"+ ($Global:ffmpegData["TrimPoints"].Count-1) +"]anull[aCut];"
        }

        function watermarkFil($op){
            if($op["Watermark Corner"] -and $op["Watermark Corner"]["Enabled"]){
                $re += "[v"+$Global:ffmpegData["fWatermark"]+"]overlay=shortest=1:" 
            
                Switch($op["Watermark Corner"]["Corner"]){
                    0{$re += "x=0:y=0"}
                    1{$re += "x=W-w:y=0"}
                    2{$re += "x=0:y=H-h"}
                    3{$re += "x=W-w:y=H-h"}
                }
            }elseif($op["Watermark Position"] -and $op["Watermark Position"]["Enabled"]){
                $re += "[v"+$Global:ffmpegData["fWatermark"]+"]overlay=shortest=1:x="+$op["Watermark Position"]["PosX"]+":y="+$op["Watermark Position"]["PosY"]
            }else{
                $re = "null"
            }
            $re
        }

        #add outro and watermark

        if($op["Outro Card"] -and $op["Outro Card"]["Enabled"]){
            if ($op["Outro Card"]["Outro Length"] -le 0){
                $null = Read-Host "Outro card length cant be zero or below."
                exit
            }
            #split cut into two parts
            $ret += "[vCut]split=2[vCut_a][vCut_b];"
            if ($op["Outro Card"]["Outro Length"] -le $op["Outro Card"]["Outro Transition"]){
                $null = Read-Host "Outro card transition exceeds outro lenght."
                exit
            }
            $cutp = $curlen-$op["Outro Card"]["Outro Length"]
            if ($cutp -le 0){
                $null = Read-Host "Outro card length exceeds trimming."
                exit
            }
            $ret += "[vCut_a]trim=end="+($cutp+$op["Outro Card"]["Outro Transition"])+",setpts=PTS-STARTPTS[vCut_a];"
            $ret += "[vCut_b]trim=start="+$cutp+",setpts=PTS-STARTPTS,scale=-1:"+($height * $op["Outro Card"]["Size"]/100)+"[vCut_b];"
            $curlen += $op["Outro Card"]["Outro Length"] + $op["Outro Card"]["Outro Transition"]
            
            #overlay second part on card
            $ret += "[v"+$Global:ffmpegData["fOutro"]+"][vCut_b]overlay=shortest=1:x="+$op["Outro Card"]["PosX"]+":y="+$op["Outro Card"]["PosY"]+"[vCut_bO];"
            #add watermark to the first part
            $ret += "[vCut_a]"+(watermarkFil $op)+"[vCut_aWm];"

            #merge two parts
            if($op["Outro Card"]["Outro Transition"] -gt 0){
                $ret += "[vCut_aWm][vCut_bO]xfade=transition=fade:duration="+$op["Outro Card"]["Outro Transition"]+":offset="+$cutp+"[vCutO];"
            }else{
                $ret += "[vCut_aWm][vCut_bO]concat[vCutO];"
            }
            $ret += "[aCut]anull[aCutO];"
        }elseif($op["Outro Video"] -and $op["Outro Video"]["Enabled"]){
            $outroLen = [double] (Invoke-Expression ($fp + " -i `"" + $op["Outro Video"]["Outro File"] + "`" -show_entries format=duration -v quiet -of csv=`"p=0 `""))
            if ($outroLen -le $op["Outro Video"]["Outro Transition"]){
                $null = Read-Host "Outro video transition can't be longer than the video length."
                exit
            }
            $ret += "[vCut]"+(watermarkFil $op)+"[vCutWm];"
            if($op["Outro Video"]["Outro Transition"] -gt 0){
                $cutp = $curlen - $op["Outro Video"]["Outro Transition"]
                $ret += "[vCutWm][v"+$Global:ffmpegData["fOutro"]+"]xfade=transition=fade:duration="+$op["Outro Video"]["Outro Transition"]+":offset="+$cutp+"[vCutO];"
                $ret += "[aCut][a"+$Global:ffmpegData["fOutro"]+"]acrossfade=d="+$op["Outro Video"]["Outro Transition"]+"[aCutO];"
            }else{
                $ret += "[vCutWm][v"+$Global:ffmpegData["fOutro"]+"]concat[vCutO];"
                $ret += "[aCut]"+"[a"+$Global:ffmpegData["fOutro"]+"]concat=v=0:a=1[aCutO];"
            }
            #we don't need this anymore
            #$curlen += $outroLen-$op["Outro Video"]["Outro Transition"]
        }else{
            $ret += "[vCut]"+(watermarkFil $op)+"[vCutO];"
            $ret += "[aCut]anull[aCutO];"
        }

        #add intro

        if($op["Intro"] -and $op["Intro"]["Enabled"]){
            $introLen = [double] (Invoke-Expression ($fp + " -i `"" + $op["Intro"]["Intro File"] + "`" -show_entries format=duration -v quiet -of csv=`"p=0 `""))
            if($op["Intro"]["Intro Transition"] -gt 0){
                $cutp = $introLen - $op["Intro"]["Intro Transition"]
                $ret += "[v"+$Global:ffmpegData["fIntro"]+"][vCutO]xfade=transition=fade:duration="+$op["Intro"]["Intro Transition"]+":offset="+$cutp+"[vCutFin];"
                $ret += "[a"+$Global:ffmpegData["fIntro"]+"][aCutO]acrossfade=d="+$op["Intro"]["Intro Transition"]+"[aCutFin];"
            }else{
                $ret += "[v"+$Global:ffmpegData["fIntro"]+"][vCutO]concat[vCutFin];"
                $ret += "[a"+$Global:ffmpegData["fIntro"]+"][aCutO]concat=v=0:a=1[aCutFin];"
            }
        }else{
            $ret += "[vCutO]null[vCutFin];"
            $ret += "[aCutO]asetpts=PTS-STARTPTS[aCutFin];"
        }

        $ret += "[vCutFin][aCutFin]concat=n=1:v=1:a=1"
        $ret
    }

    $ff += " -filter_complex `"" +(ComplexFilter)+ "`" -pix_fmt `"yuv420p`" -c:v libx264 `"" + $op["Output"]["Output File"] + $op["Output"]["Output Format"] + "`""

    Invoke-Expression $ff
}

###################################

$form = GUI

if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
    ##we invoke ffmpeg
    InvokeFFMPEG
    pause
}
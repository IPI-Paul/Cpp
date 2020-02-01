function centimetersToPoints($centimeter) {
    $point = $centimeter * 28.3464567;
    return $point;
}
function getResult($idx, $rIdx) {
    if ($idx -ne 3 -or $gridCalculate.rows[$rIdx].Cells[0].Value -ne "Total" -and !$formatting) {
        $in = normVal($gridCalculate.rows[$rIdx].Cells[1].Value); 
        $out = normVal($gridCalculate.rows[$rIdx].Cells[2].Value); 
        [MultiplyBy]::multiplyBy([ref] $in, [ref] $out);
        $gridCalculate.rows[$rIdx].Cells[3].Value = $out;
        if($gridCalculate.rows[$rIdx].Cells[3].Value -ne 0) {
            updTotals;
        }
    }
}
function idxColour($col) {
    for ($i = 0; $i -lt 256; $i++) {
        $col.Items.Add($i);
    }
}
function loadDLL() {
    if ([Environment]::Is64BitProcess) {
        Add-Type -TypeDefinition @'
            using System;
            using System.Diagnostics;
            using System.Runtime.InteropServices;

            public static class MultiplyBy {
                [DllImport("C:\\Users\\Paul\\Documents\\Source Files\\dll\\multiplyBy x64.dll", SetLastError=true, CharSet=CharSet.Ansi)]
                public static extern void multiplyBy(
                    ref double num, ref double num1
                );
            }
'@;
    } else {
        Add-Type -TypeDefinition @'
            using System;
            using System.Diagnostics;
            using System.Runtime.InteropServices;

            public static class MultiplyBy {
                [DllImport("C:\\Users\\Paul\\Documents\\Source Files\\dll\\multiplyBy x86.dll", SetLastError=true, CharSet=CharSet.Ansi)]
                public static extern void multiplyBy(
                    ref double num, ref double num1
                );
            }
'@;
    }
}
function matchColour($red, $green, $blue) {
    $distances = @{};
    foreach($obj in $namedColors) {
        $val = [Math]::Abs(([Long] $red) - ([Long] $obj.R));
        $val = $val + ([Math]::Abs(([Long] $green) - ([Long] $obj.G)));
        $val = $val + ([Math]::Abs(([Long] $blue) - ([Long] $obj.B)));
        $distances.Add($obj.Name , $val);
    }
    return [String] ($distances.GetEnumerator() | Sort-Object -Property Value | Select-Object -First 1).Name;
}
function normVal($val) {
    try {
        return [Double] $val.Replace("'", "");
    } catch {
        try {
            return [Double] $val;
        } catch {
            return [Double] 0;
        }
    }
}
function resizeItems() {
    $gridCalculate.Size = New-Object System.Drawing.Size(($objForm.Width - 2), ($objForm.Height - 88));
    $lblBgHeader.Location = New-Object System.Drawing.Size((($objForm.Width - 140) - (45 * 6)),($objForm.Height - 83));
    $cboBgRed.Location = New-Object System.Drawing.Size((($objForm.Width - 60) - (45 * 6)), ($objForm.Height - 85));
    $cboBgGreen.Location = New-Object System.Drawing.Size((($objForm.Width - 60) - (45 * 5)), ($objForm.Height - 85));
    $cboBgBlue.Location = New-Object System.Drawing.Size((($objForm.Width - 60) - (45 * 4)), ($objForm.Height - 83));
    $lblFgHeader.Location = New-Object System.Drawing.Size((($objForm.Width - 60) - (45 * 3)),($objForm.Height - 85));
    $cboFgRed.Location = New-Object System.Drawing.Size((($objForm.Width - 30) - (45 * 3)), ($objForm.Height - 85));
    $cboFgGreen.Location = New-Object System.Drawing.Size((($objForm.Width - 30) - (45 * 2)), ($objForm.Height - 85));
    $cboFgBlue.Location = New-Object System.Drawing.Size((($objForm.Width - 30) - 45), ($objForm.Height - 85));
    $cboFormats.Location = New-Object System.Drawing.Size(($objForm.Width - 332), ($objForm.Height - 62));
    $cboFunctions.Location = New-Object System.Drawing.Size(($objForm.Width - 230), ($objForm.Height - 62));
}
function restyle($typ, $val) {
    if ($cboFormats.selectedIndex -eq 0) {
        return (normVal -val $val);
    } elseif ($typ -eq 1) {
        if ($cboFormats.selectedIndex -eq 1) {
            $frmt = $cboFormats.Text.Replace("0.00", "0.####################");
        } else {
            $frmt = $cboFormats.Text.Replace("0", "0.####################");
        }
        return "{0:$frmt}" -f (normVal -val $val);
    } elseIf ($typ -eq 2) {
        if ($cboFormats.selectedIndex -eq 1) {
            $dec = 2;
        } else {
            $dec = 0;
        }
        $frmt = $cboFormats.Text;
        return "{0:$frmt}" -f ([Math]::Round((normVal -val $val), $dec));
    }
}
function rgb($red, $green, $blue) {
    $col = (([Long] $red) + (([Long] $green) * 256) + (([Long] $blue) * 65536));
    return $col;
}
function runFunction() {
    if ($cboFunctions.SelectedIndex -gt 0) {
        if ($cboFunctions.Text -eq "Send To New Email") {
            sendToNewEmail;
        } elseif ($cboFunctions.Text -eq "Send To Open Email") {
            sendToOpenEmail;
        } elseif ($cboFunctions.Text -eq "Send To New Word Document") {
            sendToNewWord;
        } elseif ($cboFunctions.Text -eq "Send To Open Word Document") {
            sendToOpenWord;
        } 

        $cboFunctions.SelectedIndex = 0;
    }
}
function sendToNewEmail() {
    $ol = New-Object -ComObject Outlook.Application;
    $oMail = $ol.CreateItem(0);
    $oMail.Display();
    $oMail.Subject = "C++ multiplyBy PowerShell Test"
    $html = tblHTML;
    $oMail.HTMLBody = "`n`n$html";
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ol);
}
function sendToOpenEmail() {
    $ol = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application");
    $ol.ActiveInspector().WordEditor.Application.Selection = "placeHere";
    $html = tblHTML;
    $ol.ActiveInspector().CurrentItem.HTMLBody = $ol.ActiveInspector().CurrentItem.htmlBody.replace("placeHere", "$html");
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ol);
}
function sendToNewWord() {
    $wrd = New-Object -ComObject Word.Application;
    $wrd.visible = $true;
    $doc = $wrd.Documents.Add();
    $doc.PageSetup.Orientation = $wConst.wdOrientLandscape;
    $doc.PageSetup | ForEach-Object {
        $_.TopMargin = centimetersToPoints -centimeter 1.75;
        $_.LeftMargin = centimetersToPoints -centimeter 1.75;
        $_.BottomMargin = centimetersToPoints -centimeter 1.75;
        $_.RightMargin = centimetersToPoints -centimeter 1.75;
    }
    tblWord -doc $doc;
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wrd);
}
function sendToOpenWord() {
    $wrd = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application");
    $doc = $wrd.ActiveDocument;
    tblWord -doc $doc;
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wrd);
}
function tblHTML() {
    $stl = "<style>`n";
    $stl = $stl + "table {`n";
    $stl = $stl + "`tborder-collapse: collapse;`n";
	$stl = $stl + "}`n";
    $stl = $stl + "table, th, td {`n";
    $stl = $stl + "`tpadding: 0px 5px 0px 5px;`n";
    $stl = $stl + "`tborder: 1 solid black;`n";
    $stl = $stl + "`tfont-size: 1em;`n";
	$stl = $stl + "}`n";
    $stl = $stl + "td {`n";
	$stl = $stl + "`tcolor: black;`n";
    $stl = $stl + "`ttext-align: right;`n";
	$stl = $stl + "}`n";
    $stl = $stl + "th {`n";
	$stl = $stl + "`tcolor: rgb(" + $cboFgRed.Text + ", " + $cboFgGreen.Text + ", " + $cboFgBlue.Text + ");`n";
	$stl = $stl + "`tbackground-color: rgb(" + $cboBgRed.Text + ", " + $cboBgGreen.Text + ", " + $cboBgBlue.Text + ");`n";
	$stl = $stl + "}`n";
    $stl = $stl + "</style>`n";
    $tbl = "<table id=""multiplyBy"">`n";
    $tr = "";
    $trd = "`t<tr>`n";
    $tre = "`t</tr>`n";
    $thd = "`t`t<th>`n`t`t`t"
    $the = "`t`t</th>`n"
    $tdd = $thd.Replace("th", "td");
    $tde = $the.Replace("th", "td");
    $th = "";
    foreach ($col in $gridCalculate.Columns) {
        if ($col.Index -gt 0) {
            $th = $th + $thd + $col.Name + $the;
        }
        $i++;
    }
    $tr = $tr + $trd + $th + $tre;
    foreach ($row in $gridCalculate.Rows) {
        if ($row.Index -eq $gridCalculate.Rows.Count - 2) {
            $tr = $tr + $trd + "<td colspan=""3"" style=""border-left: none; border-right: none;"">  &nbsp; </td>" + $tre
        } 
        if ($row.Index -lt $gridCalculate.Rows.Count - 1) {
            $td = "";
            foreach ($cell in $row.Cells) {
                if ($cell.ColumnIndex -gt 0) {
                    $td = $td + $tdd + (restyle -typ 2 -val $cell.Value) + $tde;
                }
            }
            $tr = $tr + $trd + $td + $tre;
        }
    }
    $tbl = $tbl + $tr + "</table>";
    $html = "$stl$tbl";
    return $html;
}
function tblWord($doc) {
    $tbl = $doc.Application.Selection.Tables.Add($doc.Application.Selection.Range(), $gridCalculate.Rows.Count, $gridCalculate.Columns.Count - 1);
    $tbl | ForEach-Object { 
        $_.borders.InsideLineStyle = $wConst.wdLineStyleSingle;
        $_.borders.OutsideLineStyle = $wConst.wdLineStyleSingle;
        $_.borders.InsideColor = [Long] (rgb -red $cboBgRed.Text-green $cboBgGreen.Text-blue $cboBgBlue.Text);
        $_.borders.OutsideColor = [Long] (rgb -red $cboBgRed.Text-green $cboBgGreen.Text-blue $cboBgBlue.Text);
    }
    $pg = $tbl.Rows[1].Range.Information($wConst.wdActiveEndPageNumber);
    $i = 1;
    foreach ($row in $gridCalculate.Rows) { 
        if ($i -le $gridCalculate.Rows.Count) {
            if ($i -eq $gridCalculate.Rows.Count) {
                $tbl.Rows.Add($tbl.Rows[$i]);
                $tbl.Rows[$i].Cells | ForEach-Object {
                    $_.Borders($wConst.wdBorderLeft).LineStyle = $wConst.wdLineStyleNone;
                    $_.Borders($wConst.wdBorderRight).LineStyle = $wConst.wdLineStyleNone;
                    #$_.Borders($wConst.wdBorderVertical).LineStyle = $wConst.wdLineStyleNone;
                }
                $i++;
            }
            if ($pg -lt $tbl.Rows[$i].Range.Information($wConst.wdActiveEndPageNumber)) {
                $pg = $tbl.Rows[$i].Range.Information($wConst.wdActiveEndPageNumber);
                $tbl.Rows.Add($tbl.Rows[$i]);
                $j = 1;
                foreach ($col in $gridCalculate.Columns) {
                    if ($col.Index -gt 0) {
                        $tbl.Rows[$i].Cells[$j].range | ForEach-Object { 
                            $_.text = $col.Name;
                            $_.Shading.BackgroundPatternColor = [Long] (rgb -red $cboBgRed.Text-green $cboBgGreen.Text-blue $cboBgBlue.Text);
                            $_.Font.Color = [Long] (rgb -red $cboFgRed.Text-green $cboFgGreen.Text-blue $cboFgBlue.Text);
                            $_.Font.bold = $true;
                        }
                        $j++;
                    }
                }
                $i++;
            }
            if ($i -eq 1) {
                $j = 1;
                foreach ($col in $gridCalculate.Columns) {
                    if ($col.Index -gt 0) {
                        $tbl.Rows[$i].Cells[$j].range | ForEach-Object { 
                            $_.text = $col.Name;
                            $_.Shading.BackgroundPatternColor = [Long] (rgb -red $cboBgRed.Text-green $cboBgGreen.Text-blue $cboBgBlue.Text);
                            $_.Font.Color = [Long] (rgb -red $cboFgRed.Text-green $cboFgGreen.Text-blue $cboFgBlue.Text);
                            $_.Font.bold = $true;
                        }
                        $j++;
                    }
                }
                $i++;
            }
            $j = 1;
            foreach ($cell in $row.Cells) {
                if ($cell.ColumnIndex -gt 0) {
                    $val = restyle -typ 2 -val $cell.Value;
                    $tbl.Rows[$i].Cells[$j].range.text = "$val";
                    $tbl.Rows[$i].Cells[$j].Range.ParagraphFormat.Alignment = $wConst.wdAlignParagraphRight;
                    $j++;
                }
            }
            $i++;
        }
    }
}
function updColours() {
    $bgColour = matchColour -red $cboBgRed.Text -green $cboBgGreen.Text -blue $cboBgBlue.Text;
    $fgColour = matchColour -red $cboFgRed.Text -green $cboFgGreen.Text -blue $cboFgBlue.Text;
    if ($fgColour -eq "Transparent") {
        $fgColour = "White";
    }
    $gridCalculate.ColumnHeadersDefaultCellStyle.ForeColor = $fgColour;
    $gridCalculate.ColumnHeadersDefaultCellStyle.BackColor = $bgColour;
}
function updFormats() {
    $formatting = $true;
    foreach ($row in $gridCalculate.Rows) {
        if ($row.Cells[0].Value -ne "Total" -and $row.Index -lt $gridCalculate.Rows.Count - 1) {
            foreach ($cell in $row.Cells) {
                if ($cell.ColumnIndex -gt 0) {
                    $cell.Value = (restyle -typ 1 -val $cell.Value);
                }
            }
        }
    }
    updTotals;
    $formatting = $false;
}
function updTotals() {
    if ($gridCalculate.rows[$gridCalculate.rows.count - 1].Cells[0].Value -ne "Total") {
        foreach($row in $gridCalculate.rows) {
            if ($row.Cells[0].Value -eq "Total") {
                $gridCalculate.rows.RemoveAt($row.index);
            } else {
                $num1 = $num1 + (normVal -val $row.Cells[1].Value);
                $num2 = $num2 + (normVal -val $row.Cells[2].Value);
                $res = $res + (normVal -val $row.Cells[3].Value);
            }
        }
    } 
    $num1 = $num2 = $res = 0
    foreach($row in $gridCalculate.rows) {
        if ($row.Cells[0].Value -ne "Total") {
            $num1 = $num1 + (normVal -val $row.Cells[1].Value);
            $num2 = $num2 + (normVal -val $row.Cells[2].Value);
            $res = $res + (normVal -val $row.Cells[3].Value);
        }
    }
    $gridCalculate.Rows.Add("Total", (restyle -typ 2 -val $num1), (restyle -typ 2 -val $num2), (restyle -typ 2 -val $res));
}
function viewForm() {
    $objForm = New-Object System.Windows.Forms.Form;
    $objForm.text = "IPI Paul - C++ multipyBy";
    $objForm.Size = New-Object System.Drawing.Size(470,510);
    $objForm.StartPosition = "CenterScreen";

    $gridCalculate = New-Object System.Windows.Forms.DataGridView;
    $gridCalculate.Location = New-Object System.Drawing.Size(2, 2);
    $gridCalculate.Size = New-Object System.Drawing.Size(($objForm.Width - 2), ($objForm.Height - 88));
    $gridCalculate.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True;
    $gridCalculate.AllowDrop = $true;
    $gridCalculate.AutoSize = $false;
    $gridCalculate.AutoSizeRowsMode = "AllCells";
    $gridCalculate.AutoSizeColumnsMode = "AllCells";
    $gridCalculate.ColumnCount = 4;
    $gridCalculate.Columns[0].Name = "Row Type";
    $gridCalculate.Columns[0].ReadOnly = $true;
    $gridCalculate.Columns[1].Name = "Number1";
    $gridCalculate.Columns[2].Name = "Number2";
    $gridCalculate.Columns[3].Name = "Multiplied Result";
    $gridCalculate.Columns[3].ReadOnly = $true;
    $gridCalculate.EnableHeadersVisualStyles = $false;
    $gridCalculate.add_CellValueChanged({getResult -idx $_.ColumnIndex -rIdx $gridCalculate.CurrentRow.Index;});
    $objForm.Controls.Add($gridCalculate);

    $lblBgHeader = New-Object System.Windows.Forms.Label;
    $lblBgHeader.Location = New-Object System.Drawing.Size((($objForm.Width - 140) - (45 * 6)),($objForm.Height - 83));
    $lblBgHeader.Size = New-Object System.Drawing.Size(80,20);
    $lblBgHeader.Text = "Header Back"
    $objForm.Controls.Add($lblBgHeader);

    $cboBgRed = New-Object System.Windows.Forms.ComboBox;
    $cboBgRed.Location = New-Object System.Drawing.Size((($objForm.Width - 60) - (45 * 6)), ($objForm.Height - 85));
    $cboBgRed.Size = New-Object System.Drawing.Size(45,20);
    idxColour -col $cboBgRed;
    $cboBgRed.SelectedIndex = 0;
    $cboBgRed.Add_SelectedIndexChanged({updColours;});
    $cboBgRed.Add_TextChanged({updColours;});
    $objForm.Controls.Add($cboBgRed);

    $cboBgGreen = New-Object System.Windows.Forms.ComboBox;
    $cboBgGreen.Location = New-Object System.Drawing.Size((($objForm.Width - 60) - (45 * 5)), ($objForm.Height - 85));
    $cboBgGreen.Size = New-Object System.Drawing.Size(45,20);
    idxColour -col $cboBgGreen;
    $cboBgGreen.SelectedIndex = 0;
    $cboBgGreen.Add_SelectedIndexChanged({updColours;});
    $cboBgGreen.Add_TextChanged({updColours;});
    $objForm.Controls.Add($cboBgGreen);

    $cboBgBlue = New-Object System.Windows.Forms.ComboBox;
    $cboBgBlue.Location = New-Object System.Drawing.Size((($objForm.Width - 60) - (45 * 4)), ($objForm.Height - 85));
    $cboBgBlue.Size = New-Object System.Drawing.Size(45,20);
    idxColour -col $cboBgBlue;
    $cboBgBlue.SelectedIndex = 108;
    $cboBgBlue.Add_SelectedIndexChanged({updColours;});
    $cboBgBlue.Add_TextChanged({updColours;});
    $objForm.Controls.Add($cboBgBlue);

    $lblFgHeader = New-Object System.Windows.Forms.Label;
    $lblFgHeader.Location = New-Object System.Drawing.Size((($objForm.Width - 60) - (45 * 3)),($objForm.Height - 83));
    $lblFgHeader.Size = New-Object System.Drawing.Size(30,20);
    $lblFgHeader.Text = "Fore"
    $objForm.Controls.Add($lblFgHeader);

    $cboFgRed = New-Object System.Windows.Forms.ComboBox;
    $cboFgRed.Location = New-Object System.Drawing.Size((($objForm.Width - 30) - (45 * 3)), ($objForm.Height - 85));
    $cboFgRed.Size = New-Object System.Drawing.Size(45,20);
    idxColour -col $cboFgRed;
    $cboFgRed.SelectedIndex = 255;
    $cboFgRed.Add_SelectedIndexChanged({updColours;});
    $cboFgRed.Add_TextChanged({updColours;});
    $objForm.Controls.Add($cboFgRed);

    $cboFgGreen = New-Object System.Windows.Forms.ComboBox;
    $cboFgGreen.Location = New-Object System.Drawing.Size((($objForm.Width - 30) - (45 * 2)), ($objForm.Height - 85));
    $cboFgGreen.Size = New-Object System.Drawing.Size(45,20);
    idxColour -col $cboFgGreen;
    $cboFgGreen.SelectedIndex = 255;
    $cboFgGreen.Add_SelectedIndexChanged({updColours;});
    $cboFgGreen.Add_TextChanged({updColours;});
    $objForm.Controls.Add($cboFgGreen);

    $cboFgBlue = New-Object System.Windows.Forms.ComboBox;
    $cboFgBlue.Location = New-Object System.Drawing.Size((($objForm.Width -30) - 45), ($objForm.Height - 85));
    $cboFgBlue.Size = New-Object System.Drawing.Size(45,20);
    idxColour -col $cboFgBlue;
    $cboFgBlue.SelectedIndex = 255;
    $cboFgBlue.Add_SelectedIndexChanged({updColours;});
    $cboFgBlue.Add_TextChanged({updColours;});
    $objForm.Controls.Add($cboFgBlue);

    updColours;

    $cboFormats = New-Object System.Windows.Forms.ComboBox;
    $cboFormats.Location = New-Object System.Drawing.Size(($objForm.Width - 332), ($objForm.Height - 62));
    $cboFormats.Size = New-Object System.Drawing.Size(100,20);
    $cboFormats.Items.Add("General");
    $cboFormats.Items.Add("#,###,#0.00");
    $cboFormats.Items.Add("#,###,#0");
    $cboFormats.SelectedIndex = 0;
    $cboFormats.Add_SelectedIndexChanged({updFormats;});
    $objForm.Controls.Add($cboFormats);

    $cboFunctions = New-Object System.Windows.Forms.ComboBox;
    $cboFunctions.Location = New-Object System.Drawing.Size(($objForm.Width - 230), ($objForm.Height - 62));
    $cboFunctions.Size = New-Object System.Drawing.Size(200,20);
    $cboFunctions.Items.Add("");
    $cboFunctions.Items.Add("Send to New Email");
    $cboFunctions.Items.Add("Send to Open Email");
    $cboFunctions.Items.Add("Send to New Word Document");
    $cboFunctions.Items.Add("Send to Open Word Document");
    $cboFunctions.Add_SelectedIndexChanged({runFunction;});
    $objForm.Controls.Add($cboFunctions);

    $objForm.TopMost = $False;
    $objForm.Add_Shown({$objForm.Activate()});
    $objForm.Add_Resize({resizeItems;});
    [void]$objForm.ShowDialog();        
}
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing");
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms");
$formatting = $false;
$namedColors = [System.Drawing.Color] | Get-Member -Static -MemberType Property | foreach {[Drawing.Color]::($_.Name)}
$wConst = @{
    wdOrientLandscape = 1; wdOrientPortrait = 0; wdActiveEndPageNumber = 3; wdBorderLeft = -2; wdBorderRight = -4; wdBorderVertical = -6; wdLineStyleNone = 0; wdLineStyleSingle = 1;
    wdAlignParagraphRight = 2; wdAlignParagraphCenter = 1; wdAlignParagraphLeft = 0;
}

loadDll;
viewForm;

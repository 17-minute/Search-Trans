function Search-Trans
{
    [CmdletBinding()]
    param([Parameter(Mandatory = $true, Position = 0)][string]$directory)

    
    #表示する文字列を色分けするためのクラスを定義
    #クラスには任意の文字列内の検索文字列の前後に「#」を挿入するプロパティがある
    #define a class to specify a color for a searching string
    #it has properties that insert "#" before/after a searching string in a string
    class CellInfo
    {                                                                
        [string]$fileName
        [int]$Row
        [string]$baseText
        [string]$baseWord
        [string]$corrWord
        [string]$corrText

        [string]GetReplaceBaseC()
        { return [Regex]::Replace($this.basetext,"($($this.baseword))","#`$1#") }

        [string]GetReplaceBase()
         { return [Regex]::Replace($this.basetext,"($($this.baseword))","#`$1#", "IgnoreCase") }

         [string]GetReplaceCorrC()
         {
             return [Regex]::Replace($this.CorrText,"($($this.CorrWord))","#`$1#")
         }

         [string]GetReplaceCorr()
         {
             return [Regex]::Replace($this.CorrText,"($($this.CorrWord))","#`$1#", "IgnoreCase")
         }

         CellInfo([string]$filename, [int]$row, [string]$baseword, [string]$basetext, [string]$corrword, [string]$corrtext)
         {
             [string]$this.fileName = $filename
             [int]$this.Row = $row
             [string]$this.BaseWord = $baseword
             [string]$this.BaseText = $basetext
             [string]$this.CorrWord = $corrword
             [string]$this.CorrText = $corrtext
         }  
     }

     #############################
     #メッセージボックスの設置
     #############################
     Add-Type -AssemblyName System.Windows.Forms
     Add-Type -AssemblyName Microsoft.Visualbasic
     Add-Type -AssemblyName System.Drawing

     #フォーム・タイトル
     #form and title
     $form = New-Object System.Windows.Forms.Form 
     $form.Text = "Search System for Checking Translations"
     $form.Size = New-Object System.Drawing.Size(670,350) 
     $form.StartPosition = "CenterScreen"
 
     #ラベル1
     #label 1
     $label = New-Object System.Windows.Forms.Label
     $label.Location = New-Object System.Drawing.Point(20,10) 
     $label.Size = New-Object System.Drawing.Size(670,20) 
     $label.Text = "Enter a string to search for!"
     $form.Controls.Add($label) 

     #ラベル2
     #label 2
     $label = New-Object System.Windows.Forms.Label
     $label.Location = New-Object System.Drawing.Point(20,45) 
     $label.Size = New-Object System.Drawing.Size(800,20) 
     $label.Text = "[ Search string (1) ]"
     $form.Controls.Add($label)
 
     #テキストボックス1（検索文字列）
     #textbox 1 (searching string)
     $base = New-Object System.Windows.Forms.TextBox 
     $base.Location = New-Object System.Drawing.Point(20,70) 
     $base.Multiline = $True
     $base.AcceptsReturn = $True
     $base.AcceptsTab = $True
     $base.WordWrap = $True
     $base.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
     $base.Anchor = (([System.Windows.Forms.AnchorStyles]::Left) `
                   -bor ([System.Windows.Forms.AnchorStyles]::Top) `
                   -bor ([System.Windows.Forms.AnchorStyles]::Right) `
                   -bor ([System.Windows.Forms.AnchorStyles]::Bottom))
     $base.Size = New-Object System.Drawing.Size(610,50) 
     $form.Controls.Add($base) 

     #ラベル3
     #label 3
     $label = New-Object System.Windows.Forms.Label
     $label.Location = New-Object System.Drawing.Point(20,135)
     $label.Size = New-Object System.Drawing.Size(800,20)
     $label.Text = "[ Search string (2) (in the language other than the previous string (1), Optional) ]"
     $form.Controls.Add($label)

     #テキストボックス2（任意の検索文字列）
     #textbox 2 (searching string [optional])
     $corr = New-Object System.Windows.Forms.TextBox 
     $corr.Location = New-Object System.Drawing.Point(20,160)
     $corr.Multiline = $True
     $corr.AcceptsReturn = $True
     $corr.AcceptsTab = $True
     $corr.WordWrap = $True
     $corr.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
     $corr.Anchor = (([System.Windows.Forms.AnchorStyles]::Left) `
                   -bor ([System.Windows.Forms.AnchorStyles]::Top) `
                   -bor ([System.Windows.Forms.AnchorStyles]::Right) `
                   -bor ([System.Windows.Forms.AnchorStyles]::Bottom))
     $corr.Size = New-Object System.Drawing.Size(610,50)
     $form.Controls.Add($corr) 

     #チェックボックス1
     #checkbox 1
     [System.Windows.Forms.CheckBox]$baseistarget = [System.Windows.Forms.CheckBox]::new()
     $baseistarget.size = New-Object System.Drawing.Size(25,25)
     $baseistarget.Location = New-Object System.Drawing.Point(25,220)
     $form.Controls.Add($baseistarget)

     #チェックボックス1につけるテキスト
     #text with checkbox 1
     $checktext1 = New-Object System.Windows.Forms.Label
     $checktext1.Location = New-Object System.Drawing.Point(50,222)
     $checktext1.Size = New-Object System.Drawing.Size(420,20)
     $checktext1.Text = "Search string (1) is the target language"
     $form.Controls.Add($checktext1)

     #チェックボックス2
     #checkbox 2
     [System.Windows.Forms.CheckBox]$casesensitive = [System.Windows.Forms.CheckBox]::new()
     $casesensitive.size = New-Object System.Drawing.Size(25,25)
     $casesensitive.Location = New-Object System.Drawing.Point(25,252)
     $form.Controls.Add($casesensitive)

     #チェックボックス2につけるテキスト
     #text with checkbox 2
     $checktext2 = New-Object System.Windows.Forms.Label
     $checktext2.Location = New-Object System.Drawing.Point(50,255)
     $checktext2.Size = New-Object System.Drawing.Size(240,20)
     $checktext2.Text = "Case-sensitive"
     $form.Controls.Add($checktext2)

     #OKボタン
     #OK button
     $OKButton = New-Object System.Windows.Forms.Button
     $OKButton.Location = New-Object System.Drawing.Point(470,255)
     $OKButton.Size = New-Object System.Drawing.Size(75,30)
     $OKButton.Text = "OK"
     $OKButton.Anchor = (([System.Windows.Forms.AnchorStyles]::Right) `
                    -bor ([System.Windows.Forms.AnchorStyles]::Bottom))
     $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
     $form.AcceptButton = $OKButton
     $form.Controls.Add($OKButton)
 
     #キャンセルボタン
     #cancel button
     $CancelButton = New-Object System.Windows.Forms.Button
     $CancelButton.Location = New-Object System.Drawing.Point(555,255)
     $CancelButton.Size = New-Object System.Drawing.Size(75,30)
     $CancelButton.Text = "Cancel"
     $CancelButton.Anchor = (([System.Windows.Forms.AnchorStyles]::Right) `
                        -bor ([System.Windows.Forms.AnchorStyles]::Bottom))
     $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
     $form.CancelButton = $CancelButton
     $form.Controls.Add($CancelButton)
 
     #フォームを常に手前に表示
     #always display the form in front
     $form.Topmost = $True
 
     #フォームをアクティブにし、テキストボックス1にフォーカスを設定
     #make the form active, and focus on the textbox 1
     $form.Add_Shown({$base.Select()})
 
     #フォームを表示
     #display the form
     $result = $form.ShowDialog()

     #$baseと$corrのtextプロパティを$base、$corrに再代入
     #reassign properties of $base and $corr to $base and $corr
     $base = $base.Text
     $corr = $corr.Text


     #結果の表示
     #display the result
     if($result -eq "Cancel"){ Write-Host "" -NoNewline }
     else{
         Write-Host "`nDirectory : $directory"
         Write-Host "Search string (1) : $base"
     
         if($corr){ Write-Host "Search string (2) : $corr" }

         #検索対象となる列を決定
         #decide the column to search     
         if(($BaseIsTarget.Checkstate -eq "Checked") -eq $true){
             $range_base = "C:C"
             $row_corr = 2
         }
         else{
             $range_base = "B:B"
             $row_corr = 3
         }
           
         #表示するセルの情報を格納する変数を定義
         #define the variable to assign the cell information to desplay 
         $cellInfoObjects = $null
         $cellInfoObjects = New-Object System.Collections.ArrayList

   
         #excel操作
         #excel operation
                  
         #$baseをexcelファイル内の指定した列内を検索し、一致したら変数に格納
         #Search $base in each excel file for specified column and 
         #if there are cell texts that match $base, assign the text to the variable
         $xlapp = New-Object -ComObject Excel.Application                                                                      
         Get-ChildItem $directory -Include "*.xlsx","*.xls","*.xlsm" -Recurse -Name |                                                      
              ForEach-Object{                                           
                   $filename = $_
                   $wb = $xlapp.Workbooks.Open("$directory\$filename")
                   $ws = $xlapp.Worksheets.Item(1)
                                                                     
                   if($caseSensitive.Checkstate -eq "Checked"){
                        $match = $firstmatch = $ws.range($range_base).Find($base,[type]::Missing,[type]::Missing,[type]::Missing,[type]::Missing,[type]::Missing,$True)                                                                       
                    }
                   else{
                        $match = $firstmatch = $ws.range($range_base).Find($base)
                   }

                   While($match -ne $null){
                       [void]$cellInfoObjects.Add([CellInfo]::new("$filename", $match.row, "$base", "$($match.text)", "$corr", "$($ws.Cells.Item($match.row,$row_corr).text)"))                                                                  
                       $match = $ws.range($range_base).findnext($match)
                       if($match.Address() -eq $firstmatch.Address()){ break }
                   }                            
                   $wb.Close()                                                                      
              }

          #文字列内の「#」を区切り文字にして出力するクラスを定義
          #define the class that output a string specifying 「#」as the delimiter
          class ColoriseText{
             static[void]OutputText($string, $keyword){
                 @($string -split "#")|
                     ForEach-Object{
                         if($_ -match $keyword){
                              Write-Host $_ -ForegroundColor Red -NoNewline                                                                
                          }
                          else{
                             Write-Host $_ -NoNewline
                          }
                     }
              }

             ColoriseText([string]$string, [string]$keyword){
                 $this.string = $string
                 $this.word = $keyword
             }
          }


         
           #出力に使える文字色をランダムに選択
           #原文と訳文の区切る線を$lineに格納
           #choose a color randomly  
           #assign lines to separate between original and translation
           $colour = [enum]::GetValues($host.UI.RawUI.ForegroundColor.GetType())| Get-Random                                                     
           $line = "`n==========================================================="

           #-corrパラメータの引数に一致するものがなかった時にセルの情報を格納する変数を定義
           #define the variable to assign cell information when there is no argument in -corr parameter
           $notMatch = $null
           $notMatch = New-Object System.Collections.ArrayList



        #  出力
        
        Write-Host "`n`n`n$($cellInfoObjects.Count) lines have been found in total!" -ForegroundColor Cyan

        #常に「原文-訳文」の順に出力
        #検索文字列がセルのテキスト内にある場合、その箇所だけ文字色を赤にして出力
        #セルのテキスト内に任意の検索文字列がない場合は$notMatchに格納して別に出力
        #always output the order "original-translation"
        #Output the searching string with the letters red if the cell text includes it 
        #if the cell text doesn't include the searching string, assign the cell text to $notMatch and output separately
        
          if($baseistarget.Checkstate -eq "Checked"){
               ForEach($cellInfoObject in $cellInfoObjects){                               
                   if(-not $corr){                         
                        Write-Host "`n`n`n$($cellInfoObject.filename), line $($cellInfoObject.row)" -ForegroundColor Cyan
                        Write-Host $cellInfoObject.CorrText -NoNewline                                               
                        Write-Host $line -ForegroundColor $colour
                        [ColoriseText]::OutputText($cellInfoObject.GetReplaceBase(),$base)
                   }
                   else{  
                       if($CaseSensitive.CheckState -eq "Checked"){
                           if($cellInfoObject.GetReplaceCorrC() -match "#"){
                                Write-Host "`n`n`n$($cellInfoObject.filename), line $($cellInfoObject.row)" -ForegroundColor Cyan                                                     
                                [Colorisetext]::OutputText($cellInfoObject.GetReplaceCorrC(), $corr)
                                Write-Host $line -ForegroundColor $colour
                                [ColoriseText]::OutputText($cellInfoObject.GetReplaceBaseC(), $base)
                            }
                            else{
                                [void]$notMatch.Add($cellInfoObject)
                            }
                        }
                        elseif($CaseSensitive.CheckState -ne "Checked"){
                            if($cellInfoObject.GetReplaceCorr() -match "#"){
                                 Write-Host "`n`n`n$($cellInfoObject.filename), line $($cellInfoObject.row)" -ForegroundColor Cyan                                                     
                                 [Colorisetext]::OutputText($cellInfoObject.GetReplaceCorr(), $corr)
                                 Write-Host $line -ForegroundColor $colour
                                 [ColoriseText]::OutputText($cellInfoObject.GetReplaceBase(), $base)
                            }
                            else{
                                [void]$notMatch.Add($cellInfoObject)
                            }
                       }                       
                    }                                               
               }

               if($notMatch){
                   Write-Host "`n`n`n$($NotMatch.Count) lines that don't include the search string (2) have been found!" -ForegroundColor Yellow -NoNewline
                   ForEach($nm in $notMatch){                              
                       Write-Host "`n`n`n$($nm.filename), line $($nm.row)" -ForegroundColor Yellow
                       Write-Host $nm.corrText -NoNewline
                       Write-Host $line -ForegroundColor $colour
                       [ColoriseText]::OutputText($nm.GetReplaceBase(), $base)                                          
                   }
               }                    
           }                 
           else{
               ForEach($cellInfoObject in $cellInfoObjects){                        
                   if(-not $corr){                          
                        Write-Host "`n`n$($cellInfoObject.filename), line $($cellInfoObject.row)" -ForegroundColor Cyan
                        [ColoriseText]::OutputText($cellInfoObject.GetReplaceBase(), $base)
                        Write-Host $line -ForegroundColor $colour
                        Write-Host $cellInfoObject.CorrText                                               
                   }
                   else{                                          
                       if($CaseSensitive.CheckState -eq "Checked"){
                            if($cellInfoObject.GetReplaceCorrC() -match "#"){                                                        
                                 Write-Host "`n`nline $($cellInfoObject.filename), line $($cellInfoObject.row)" -ForegroundColor Cyan
                                 [ColoriseText]::OutputText($cellInfoObject.GetReplaceBaseC(), $base)
                                 Write-Host $line -ForegroundColor $colour
                                 [ColoriseText]::OutputText($cellInfoObject.GetReplaceCorrC(), $corr)                                  
                            }
                            else{
                                 [void]$notMatch.Add($cellInfoObject)
                            }
                        }
                        elseif($CaseSensitive.CheckState -eq "Unchecked"){
                            if($cellInfoObject.GetReplaceCorr() -match "#"){
                                 Write-Host "`n`n$($cellInfoObject.filename), line $($cellInfoObject.row)" -ForegroundColor Cyan
                                 [ColoriseText]::OutputText($cellInfoObject.GetReplaceBase(), $base)
                                 Write-Host $line -ForegroundColor $colour
                                 [ColoriseText]::OutputText($cellInfoObject.GetReplaceCorr(), $corr)
                            }
                            else{                                                         
                                 [void]$notMatch.Add($cellInfoObject)
                            }                               
                        }                                                
                  }
                }
                                        
           if($notMatch){        
                Write-Host "`n`n`n$($NotMatch.Count) lines that don't include the search string (2) have been found!" -ForegroundColor Yellow -NoNewline

                ForEach($nm in $notMatch){                             
                     Write-Host "`n`n$($nm.filename), line $($nm.row)" -ForegroundColor Yellow
                     [ColoriseText]::OutputText($nm.GetReplaceBase(), $base)
                     Write-Host $line -ForegroundColor $colour
                     Write-Host $nm.CorrText                          
                }
          }
           }
                 
       $xlapp.Quit()
       $ws,$wb,$xlapp|
       ForEach-Object{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($_)}
    }
}

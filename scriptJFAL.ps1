function EnviarMensagem($secao) {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  

    $res = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds

    $Form = New-Object System.Windows.Forms.Form
    $form.minimizebox = $true
    $form.maximizebox = $true
    $Form.Size = New-Object System.Drawing.Size(440, 450)
    $Form.Text = "NTI-JFAL"
    $Form.SizeGripStyle = "Hide"
    $Form.ShowInTaskbar = $False  
    $Form.StartPosition = "CenterScreen"
    
    ##
    $Font = New-Object System.Drawing.Font("IMPACT", 30, [System.Drawing.FontStyle]::Italic)

    $Titulo = New-Object System.Windows.Forms.Label
    $Titulo.Text = "ENTRE COM A MENSAGEM"
    $Titulo.Height = (50)
    $Titulo.Width = (500)
    $Titulo.Font = $Font
    $Form.Controls.Add($Titulo)

    $caixademensagem = New-Object System.Windows.Forms.TextBox 
    $caixademensagem.Location = New-Object System.Drawing.Size(10, 60) 
    $caixademensagem.Size = New-Object System.Drawing.Size(400, 150) 
    $caixademensagem.MultiLine = $True 
    $caixademensagem.ScrollBars = "Vertical" 
    $Form.Controls.Add($caixademensagem)
    #####
    $outputBox = New-Object system.windows.Forms.TextBox
    $outputBox.Multiline = $true
    $outputBox.Width = 400
    $outputBox.Height = 150
    $outputBox.ReadOnly = $true
    $outputBox.ScrollBars = "Vertical"
    $outputBox.WordWrap = $false
    $outputBox.location = new-object system.drawing.point(10, 250)
    $outputBox.Font = "Lucida Sans ,9"
    $Form.Controls.Add($outputBox)
    #####
    $Botao_Enviar = New-Object System.Windows.Forms.Button
    $Botao_Enviar.Width = 80
    $Botao_Enviar.Height = 20
    $Botao_Enviar.Text = ("Enviar")
    $Botao_Enviar.Location = New-Object System.Drawing.Size(330, 220)
    $Botao_Enviar.Add_Click( {
            $outputBox.Appendtext("Unidade escolhida: $secao`n")
            $outputBox.Appendtext("`n")
            $mensagem = $caixademensagem.Text
            Get-ADComputer -Filter * -SearchBase "OU=$Secao, OU=JFAL-Computadores, dc=jfal, dc=jus, dc=br" | ForEach-Object {
                $computador = $_.Name
                $outputBox.Appendtext("Enviando mensagem para: $computador - Conteudo: $mensagem `n")       
                if (testconnection($computador)) {
                    msg /server:$computador * /time:3600 "$mensagem"
                    $outputBox.Appendtext("ENVIADO!`n")    
                    $outputBox.Appendtext("`n")
                }
                else {
                    $outputBox.Appendtext("$computador Desconectado!`n") 
                    $outputBox.Appendtext("`n")   
                }            
            }
            $outputBox.Appendtext("`n")
            $outputBox.Appendtext("Tarefa Concluida!`n")
        
        }) 
    $Form.Controls.Add($Botao_Enviar)
   
    ##

    $Botao_Fechar = New-Object System.Windows.Forms.Button
    $Botao_Fechar.Width = 80
    $Botao_Fechar.Height = 20
    $Botao_Fechar.Text = ("Fechar")
    $Botao_Fechar.Location = New-Object System.Drawing.Size(330, 410)
    $Botao_Fechar.Add_Click( {$Form.Close()})
    $Form.Controls.Add($Botao_Fechar)


    $result = $form.ShowDialog()

    $Form.Add_Shown( {$Form.Activate()})
    
}

function GPUpdate($Secao) { 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  

    $res = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds

    $Form = New-Object System.Windows.Forms.Form
    $form.minimizebox = $true
    $form.maximizebox = $true
    $Form.Size = New-Object System.Drawing.Size(480, 380)
    $Form.Text = "NTI-JFAL"
    $Form.SizeGripStyle = "Hide"
    $Form.ShowInTaskbar = $False  
    $Form.StartPosition = "CenterScreen"
    
    $Titulo = New-Object System.Windows.Forms.Label
    $Titulo.Text = "GPUpdate - Atualizar Diretivas de Grupos"
    $Titulo.Height = (50)
    $Titulo.Width = (500)
    $Titulo.Font = $Font
    $Form.Controls.Add($Titulo)

    $outputBox = New-Object system.windows.Forms.TextBox
    $outputBox.Multiline = $true
    $outputBox.Width = 400
    $outputBox.Height = 150
    $outputBox.ReadOnly = $true
    $outputBox.ScrollBars = "Vertical"
    $outputBox.WordWrap = $false
    $outputBox.location = new-object system.drawing.point(20, 150)
    $outputBox.Font = "Lucida Sans ,9"
    $Form.Controls.Add($outputBox)

    $Botao_Fechar = New-Object System.Windows.Forms.Button
    $Botao_Fechar.Width = 80
    $Botao_Fechar.Height = 20
    $Botao_Fechar.Text = ("Fechar")
    $Botao_Fechar.Location = New-Object System.Drawing.Size(250, 310)
    $Botao_Fechar.Add_Click( {$Form.Close()})
    $Form.Controls.Add($Botao_Fechar)

    $CheckBox_Force = New-Object System.Windows.Forms.CheckBox
    $CheckBox_Force.Width = 500
    $CheckBox_Force.Height = 30
    $CheckBox_Force.Text = "FORCE - Força o comando a ser executado sem solicitar a confirmação do usuário."
    $CheckBox_Force.Location = New-Object System.Drawing.Size(20, 50)
    $Form.Controls.Add($CheckBox_Force)

    $CheckBox_Boot = New-Object System.Windows.Forms.CheckBox
    $CheckBox_Boot.Width = 470
    $CheckBox_Boot.Height = 50
    $CheckBox_Boot.Text = "BOOT - Indica que o cmdlet reinicia o computador depois que as configurações da Diretiva de Grupo são aplicadas. "
    $CheckBox_Boot.Location = New-Object System.Drawing.Size(20, 90)
    $Form.Controls.Add($CheckBox_Boot)
    
    $Botao_Confirmar = New-Object System.Windows.Forms.Button
    $Botao_Confirmar.Width = 80
    $Botao_Confirmar.Height = 20
    $Botao_Confirmar.Text = ("Confirmar")
    $Botao_Confirmar.Location = New-Object System.Drawing.Size(340, 310)
    $Botao_Confirmar.Add_Click( {
            if ($CheckBox_Force.Checked -eq $true -and $CheckBox_Boot.Checked -eq $true) {
                $ErrorActionPreference = "Silentilycontinue"
                $outputBox.Appendtext("Unidade escolhida: $secao`n")
                $outputBox.AppendText("`n") 
                Get-ADComputer -Filter * -SearchBase "OU=$Secao, OU=JFAL-Computadores, dc=jfal, dc=jus, dc=br" | ForEach-Object {
                    $computador = $_.Name  
                    $outputBox.Appendtext("Atualizando as diretivas de grupos de $computador`n")        
                    if (testconnection($computador)) {
                        Invoke-GPUpdate –computer $computador -force -Boot
                        $outputBox.AppendText( "$computador Atualizado!`n")
                        $outputBox.AppendText("`n")
                    }
                    else {
                        $outputBox.AppendText( "$computador Desconectado!`n")
                        $outputBox.AppendText("`n")
                    }
    
                }
            }
            elseif ($CheckBox_Force.Checked -eq $true -and $CheckBox_Boot.Checked -eq $false) {
                $ErrorActionPreference = "Silentilycontinue"
                $outputBox.Appendtext("Unidade escolhida: $secao`n")
                $outputBox.AppendText("`n") 
                Get-ADComputer -Filter * -SearchBase "OU=$Secao, OU=JFAL-Computadores, dc=jfal, dc=jus, dc=br" | ForEach-Object {
                    $computador = $_.Name  
                    $outputBox.Appendtext("Atualizando as diretivas de grupos de $computador`n")        
                    if (testconnection($computador)) {
                        Invoke-GPUpdate –computer $computador -force
                        $outputBox.AppendText( "$computador Atualizado!`n")
                        $outputBox.AppendText("`n")
                    }
                    else {
                        $outputBox.AppendText( "$computador Desconectado!`n")
                        $outputBox.AppendText("`n")
                    }
    
                }
            }
            elseif ($CheckBox_Force.Checked -eq $false -and $CheckBox_Boot.Checked -eq $true) {
                $ErrorActionPreference = "Silentilycontinue"
                $outputBox.Appendtext("Unidade escolhida: $secao`n")
                $outputBox.AppendText("`n") 
                Get-ADComputer -Filter * -SearchBase "OU=$Secao, OU=JFAL-Computadores, dc=jfal, dc=jus, dc=br" | ForEach-Object {
                    $computador = $_.Name  
                    $outputBox.Appendtext("Atualizando as diretivas de grupos de $computador`n")        
                    if (testconnection($computador)) {
                        Invoke-GPUpdate –computer $computador -Boot
                        $outputBox.AppendText( "$computador Atualizado!`n")
                        $outputBox.AppendText("`n")
                    }
                    else {
                        $outputBox.AppendText( "$computador Desconectado!`n")
                        $outputBox.AppendText("`n")
                    }
    
                }
            }
            elseif ($CheckBox_Force.Checked -eq $false -and $CheckBox_Boot.Checked -eq $false) {
                $ErrorActionPreference = "Silentilycontinue"
                $outputBox.Appendtext("Unidade escolhida: $secao`n")
                $outputBox.AppendText("`n") 
                Get-ADComputer -Filter * -SearchBase "OU=$Secao, OU=JFAL-Computadores, dc=jfal, dc=jus, dc=br" | ForEach-Object {
                    $computador = $_.Name  
                    $outputBox.Appendtext("Atualizando as diretivas de grupos de $computador`n")        
                    if (testconnection($computador)) {
                        Invoke-GPUpdate –computer $computador 
                        $outputBox.AppendText( "$computador Atualizado!`n")
                        $outputBox.AppendText("`n")
                    }
                    else {
                        $outputBox.AppendText( "$computador Desconectado!`n")
                        $outputBox.AppendText("`n")
                    }
    
                }
            }
            $outputBox.AppendText("Tarefa Concluida!") 
        })
    $Form.Controls.Add($Botao_Confirmar)
       
    $result = $form.ShowDialog()

    $Form.Add_Shown( {$Form.Activate()})
}
    
function testconnection($nomecomputador) {
        Test-Connection $nomecomputador -Quiet -Count 1            
    }

function UserName($computador){
    $CS = Gwmi Win32_ComputerSystem -Comp "$computador"
    
    return $CS.UserName
}

function ErroSecao{
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  

        $res = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds

        $Form_selecione = New-Object System.Windows.Forms.Form
        $form_selecione.minimizebox = $true
        $form_selecione.maximizebox = $true
        $Form_selecione.Size = New-Object System.Drawing.Size(250, 100)
        $Form_selecione.Text = "ERRO"
        $Form_selecione.SizeGripStyle = "Hide"
        $Form_selecione.ShowInTaskbar = $False  
        $Form_selecione.StartPosition = "CenterScreen"

        $Titulo = New-Object System.Windows.Forms.Label
        $Titulo.Text = "Selecione uma seção"
        $Titulo.Height = (50)
        $Titulo.Width = (120)
        $Titulo.Location = New-Object System.Drawing.Size(30, 20)
        $Form_selecione.Controls.Add($Titulo)
    
        $Button_ok = New-Object System.Windows.Forms.Button 
        $Button_ok.Location = New-Object System.Drawing.Size(160, 10) 
        $Button_ok.Size = New-Object System.Drawing.Size(30, 30) 
        $Button_ok.text = "OK"
        $Button_ok.Add_Click( {$Form_selecione.Close()})
        $Form_selecione.Controls.Add($Button_ok)
    
    
    
    
    
        $result = $form_selecione.ShowDialog()

        $Form_selecione.Add_Shown( {$Form_selecione.Activate()})

    }

function ErroComputador {
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  

        $res = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds

        $Form_selecione = New-Object System.Windows.Forms.Form
        $form_selecione.minimizebox = $true
        $form_selecione.maximizebox = $true
        $Form_selecione.Size = New-Object System.Drawing.Size(250, 100)
        $Form_selecione.Text = "ERRO"
        $Form_selecione.SizeGripStyle = "Hide"
        $Form_selecione.ShowInTaskbar = $False  
        $Form_selecione.StartPosition = "CenterScreen"

        $Titulo = New-Object System.Windows.Forms.Label
        $Titulo.Text = "Selecione um computador"
        $Titulo.Height = (50)
        $Titulo.Width = (120)
        $Titulo.Location = New-Object System.Drawing.Size(30, 20)
        $Form_selecione.Controls.Add($Titulo)
    
        $Button_ok = New-Object System.Windows.Forms.Button 
        $Button_ok.Location = New-Object System.Drawing.Size(160, 10) 
        $Button_ok.Size = New-Object System.Drawing.Size(30, 30) 
        $Button_ok.text = "OK"
        $Button_ok.Add_Click( {$Form_selecione.Close()})
        $Form_selecione.Controls.Add($Button_ok)
    
    
    
    
    
        $result = $form_selecione.ShowDialog()

        $Form_selecione.Add_Shown( {$Form_selecione.Activate()})

    }

function GPResult_mensagens($mensagem){
            [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
            [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  

            $res = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds

            $Form_selecione = New-Object System.Windows.Forms.Form
            $form_selecione.minimizebox = $true
            $form_selecione.maximizebox = $true
            $Form_selecione.Size = New-Object System.Drawing.Size(390, 100)
            $Form_selecione.Text = "NTI-JFAL"
            $Form_selecione.SizeGripStyle = "Hide"
            $Form_selecione.ShowInTaskbar = $False  
            $Form_selecione.StartPosition = "CenterScreen"

            $Titulo = New-Object System.Windows.Forms.Label
            $Titulo.Text = $mensagem
            $Titulo.Height = (50)
            $Titulo.Width = (300)
            $Titulo.Location = New-Object System.Drawing.Size(20, 20)
            $Form_selecione.Controls.Add($Titulo)
    
            $Button_ok = New-Object System.Windows.Forms.Button 
            $Button_ok.Location = New-Object System.Drawing.Size(330,20) 
            $Button_ok.Size = New-Object System.Drawing.Size(30, 30) 
            $Button_ok.text = "OK"
            $Button_ok.Add_Click( {$Form_selecione.Close()})
            $Form_selecione.Controls.Add($Button_ok)
    
    
    
    
    
            $result = $form_selecione.ShowDialog()

            $Form_selecione.Add_Shown( {$Form_selecione.Activate()})
}

function GPResult($computador){
     if(testconnection($computador)){   
         $usuario=UserName($computador)
         if($usuario -eq $null){
            $mensagem="Nenhum usuario logado!"
            GPResult_mensagens($mensagem)
         }else{
            Get-GPResultantSetOfPolicy -Computer $computador -User $usuario -Path $env:userprofile\Desktop\GPResult.html -ReportType html
            $mensagem="Tarefa concluida com sucesso!`nDisponível em $env:userprofile\Desktop\GPResult.html"
            GPResult_mensagens($mensagem)        
         }
    }
    else{
        $mensagem="Computador Desconectado!"
        GPResult_mensagens($mensagem)
    }
}

function Menu {
        do {
            write-host ""
            write-host "############ MENU #############"
            write-host ""
            write-host "1-Enviar Mensagem"
            write-host "2-Atualizar Diretivas de Grupos"
            write-host "0-Sair"

            $opcao = read-host "Digite o numero da opção"

            switch ($opcao) {
                1 {EnviarMensagem}
                2 {GPUpdate}
                0 {return}            
            }
            pause
        }until ($opcao -eq '0')
    }

function startmenu {
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  

        $res = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds

        $Form = New-Object System.Windows.Forms.Form
        $form.minimizebox = $true
        $form.maximizebox = $true
        $Form.Size = New-Object System.Drawing.Size(650, 500)
        $Form.Text = "NTI-JFAL"
        $Form.SizeGripStyle = "Hide"
        $Form.ShowInTaskbar = $False  
        $Form.StartPosition = "CenterScreen"


        $Font = New-Object System.Drawing.Font("IMPACT", 20, [System.Drawing.FontStyle]::Italic)

        $ListBox_secoes = New-Object System.Windows.Forms.ListBox
        $ListBox_secoes.Name = "Seções"
        $ListBox_secoes.Width = 200
        $ListBox_secoes.Height = 400
        $ListBox_secoes.Location = New-Object System.Drawing.Size(40, 20)    
        $Form.Controls.Add($ListBox_secoes)

        $ListBox_pcs = New-Object System.Windows.Forms.ListBox
        $ListBox_pcs.Name = "Seções"
        $ListBox_pcs.Width = 200
        $ListBox_pcs.Height = 400
        $ListBox_pcs.Location = New-Object System.Drawing.Size(250, 20)    
        $Form.Controls.Add($ListBox_pcs)

        $Button_EnviarMensagem = New-Object System.Windows.Forms.Button 
        $Button_EnviarMensagem.Location = New-Object System.Drawing.Size(470, 20) 
        $Button_EnviarMensagem.Size = New-Object System.Drawing.Size(150, 80) 
        $Button_EnviarMensagem.font = $font
        $Button_EnviarMensagem.forecolor = "Black"
        $Button_EnviarMensagem.text = "Enviar Mensagem"
        $Button_EnviarMensagem.Add_Click( {
                if ($ListBox_secoes.SelectedIndices -ne $null) {
                    EnviarMensagem($ListBox_secoes.SelectedItems)
                }
                else {
                    errosecao
                }
            }) 
        $Form.Controls.Add($Button_EnviarMensagem) 

        $Button_GPUpdate = New-Object System.Windows.Forms.Button 
        $Button_GPUpdate.Location = New-Object System.Drawing.Size(470, 110) 
        $Button_GPUpdate.Size = New-Object System.Drawing.Size(150, 80) 
        $Button_GPUpdate.font = $font
        $Button_GPUpdate.forecolor = "Black"
        $Button_GPUpdate.text = "GPUpdate"
        $Button_GPUpdate.Add_Click( {
                if ($ListBox_secoes.SelectedIndices -ne $null) {
                    GPUpdate($ListBox_secoes.SelectedItems)
                }
                else {
                    errosecao
                }     
            }) 
        $Form.Controls.Add($Button_GPUpdate) 



        $Button_GPResult = New-Object System.Windows.Forms.Button 
        $Button_GPResult.Location = New-Object System.Drawing.Size(470, 210) 
        $Button_GPResult.Size = New-Object System.Drawing.Size(150, 80) 
        $button_GPResult.font = $font
        $button_GPResult.forecolor = "Black"
        $button_GPResult.text = "GPResult"
        $Button_GPResult.Add_Click( {
                if ($ListBox_pcs.SelectedIndices -ne $null){
                    GPResult($ListBox_pcs.SelectedItem)
                }
                else {
                    ErroComputador
                }       
            }) 
        $Form.Controls.Add($Button_GPResult)   



        $Button_Cancelar = New-Object System.Windows.Forms.Button 
        $Button_Cancelar.Location = New-Object System.Drawing.Size(470, 410) 
        $Button_Cancelar.Size = New-Object System.Drawing.Size(150, 30) 
        $button_Cancelar.font = $font
        $button_Cancelar.forecolor = "Black"
        $button_Cancelar.text = "Cancelar"
        $Button_Cancelar.Add_Click( {$Form.Close()}) 
        $Form.Controls.Add($Button_Cancelar)

        $Button_Computadores = New-Object System.Windows.Forms.Button 
        $Button_Computadores.Location = New-Object System.Drawing.Size(140, 415) 
        $Button_Computadores.Size = New-Object System.Drawing.Size(100, 20) 
        $button_Computadores.forecolor = "Black"
        $button_Computadores.text = "Computadores"
        $Button_Computadores.Add_Click( {
                if ($ListBox_secoes.SelectedIndices -ne $null) {
                    $secao = $ListBox_secoes.SelectedItems
                    $ListBox_pcs.Items.Clear()
                    $ListBox_pcs.Items.Add("COMPUTADORES")
                    Get-ADComputer -Filter * -SearchBase "OU=$secao, OU=JFAL-Computadores, dc=jfal, dc=jus, dc=br" | ForEach-Object {
                        $computador = $_.Name
                        $ListBox_pcs.Items.Add($computador)
                    }
                }
                else {
                    errosecao
                } 
            }) 
        $Form.Controls.Add($Button_Computadores)

        $Button_secoes = New-Object System.Windows.Forms.Button 
        $Button_secoes.Location = New-Object System.Drawing.Size(40, 415) 
        $Button_secoes.Size = New-Object System.Drawing.Size(100, 20) 
        $button_secoes.forecolor = "Black"
        $button_secoes.text = "Seções"
        $Button_secoes.Add_Click( {      
                $ListBox_secoes.Items.Clear()
                Get-ADOrganizationalUnit -Filter * -SearchBase "OU=JFAL-Computadores, dc=jfal, dc=jus, dc=br" | ForEach-Object {
                    $secoes = $_.Name
                    $ListBox_secoes.Items.Add("$secoes")
                }
            }) 
        $Form.Controls.Add($Button_secoes)

        $result = $form.ShowDialog()

        $Form.Add_Shown( {$Form.Activate()})

    }

startmenu

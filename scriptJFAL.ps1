

#ENVIA MENSAGEM VIA msg.exe PARA UMA SEÇÃO
function EnviarMensagem($secao) {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  

    #CONFIGURA A JANELA
    $Form = New-Object System.Windows.Forms.Form
    $form.minimizebox = $false
    $form.maximizebox = $false
    $Form.Size = New-Object System.Drawing.Size(440, 475)
    $Form.Text = "ENVIAR MENSAGEM"
    $Form.StartPosition = "CenterScreen"
    
    #TITULO DA TELA
    $Titulo = New-Object System.Windows.Forms.Label
    $Titulo.Text = "ENTRE COM A MENSAGEM"
    $Titulo.Size = New-Object System.Drawing.Size(500, 50)
    $Titulo.Font = New-Object System.Drawing.Font("IMPACT", 30, [System.Drawing.FontStyle]::Italic)
    $Form.Controls.Add($Titulo)

    #CAIXA PARA DIGITAR A MENSAGEM
    $caixademensagem = New-Object System.Windows.Forms.TextBox 
    $caixademensagem.Location = New-Object System.Drawing.Size(10, 60) 
    $caixademensagem.Size = New-Object System.Drawing.Size(400, 150) 
    $caixademensagem.MultiLine = $True 
    $caixademensagem.ScrollBars = "Vertical" 
    $Form.Controls.Add($caixademensagem)
   
    #LOG DE SAIDA PARA USUARIO 
    $outputBox = New-Object system.windows.Forms.TextBox
    $outputBox.Multiline = $true
    $outputBox.Size = New-Object System.Drawing.Size(400, 150)
    $outputBox.ReadOnly = $true
    $outputBox.ScrollBars = "Vertical"
    $outputBox.location = new-object system.drawing.point(10, 250)
    $outputBox.Font = "Lucida Sans,9"
    $Form.Controls.Add($outputBox)
        
    #BOTÃO PARA ENVIAR MENSAGEM
    $Botao_Enviar = New-Object System.Windows.Forms.Button
    $Botao_Enviar.Size = New-Object System.Drawing.Size(80, 20)
    $Botao_Enviar.Text = ("Enviar")
    $Botao_Enviar.Location = New-Object System.Drawing.Size(330, 220)
    $Botao_Enviar.Add_Click( {
            #ALGORITIMO PARA ENVIAR MENSAGEM
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
   
   
    #FECHAR JANELA
    $Botao_Fechar = New-Object System.Windows.Forms.Button
    $Botao_Fechar.Size= New-Object System.Drawing.Size(80, 20)
    $Botao_Fechar.Text = ("Fechar")
    $Botao_Fechar.Location = New-Object System.Drawing.Size(330, 410)
    $Botao_Fechar.Add_Click({$Form.Close()})
    $Form.Controls.Add($Botao_Fechar)

    #INICIA A JANELA
    $result = $form.ShowDialog()
    
    
}

#EXECUTA COMANDO GPUPDATE EM UMA SEÇÃO
function GPUpdate($Secao) { 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  
    
    #CONFIGURA JANELA
    $Form = New-Object System.Windows.Forms.Form
    $form.minimizebox = $False
    $form.maximizebox = $False
    $Form.Size = New-Object System.Drawing.Size(480, 380)
    $Form.Text = "GPUpdate"
    $Form.StartPosition = "CenterScreen"
    
    #TITULO DA TELA
    $Titulo = New-Object System.Windows.Forms.Label
    $Titulo.Text = "GPUpdate - Atualizar Diretivas de Grupos"
    $Titulo.Size = New-Object System.Drawing.Size(500, 50)
    $Titulo.Font = New-Object System.Drawing.Font("IMPACT", 20, [System.Drawing.FontStyle]::Italic)
    $Form.Controls.Add($Titulo)

    #LOG DE SAIDA PARA USUARIO
    $outputBox = New-Object system.windows.Forms.TextBox
    $outputBox.Multiline = $true
    $outputBox.Size = New-Object System.Drawing.Size(400, 150)
    $outputBox.ReadOnly = $true
    $outputBox.ScrollBars = "Vertical"
    $outputBox.location = new-object system.drawing.point(20, 150)
    $outputBox.Font = "Lucida Sans ,9"
    $Form.Controls.Add($outputBox)

    #BOTAO PARA FECHAR JANELA
    $Botao_Fechar = New-Object System.Windows.Forms.Button
    $Botao_Fechar.Size = New-Object System.Drawing.Size(80, 20)
    $Botao_Fechar.Text = ("Fechar")
    $Botao_Fechar.Location = New-Object System.Drawing.Size(250, 310)
    $Botao_Fechar.Add_Click( {$Form.Close()})
    $Form.Controls.Add($Botao_Fechar)

    #CHECAGEM DO PARAMENTRO -FORCE
    $CheckBox_Force = New-Object System.Windows.Forms.CheckBox
    $CheckBox_Force.Size = New-Object System.Drawing.Size(500, 30)
    $CheckBox_Force.Text = "FORCE - Força o comando a ser executado sem solicitar a confirmação do usuário."
    $CheckBox_Force.Location = New-Object System.Drawing.Size(20, 50)
    $Form.Controls.Add($CheckBox_Force)

    #CHECAGEM DO PARAMETRO -BOOT
    $CheckBox_Boot = New-Object System.Windows.Forms.CheckBox
    $CheckBox_Boot.Size = New-Object System.Drawing.Size(470, 50)
    $CheckBox_Boot.Text = "BOOT - Indica que o cmdlet reinicia o computador depois que as configurações da Diretiva de Grupo são aplicadas. "
    $CheckBox_Boot.Location = New-Object System.Drawing.Size(20, 90)
    $Form.Controls.Add($CheckBox_Boot)
    
    #EXECULTA O GPUpdate
    $Botao_Confirmar = New-Object System.Windows.Forms.Button
    $Botao_Confirmar.Size = New-Object System.Drawing.Size(80, 20)
    $Botao_Confirmar.Text = ("Confirmar")
    $Botao_Confirmar.Location = New-Object System.Drawing.Size(340, 310)
    $Botao_Confirmar.Add_Click( {
            #ALGORITIMO DO GPUpdate
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
    
    #INICIA A JANELA   
    $result = $form.ShowDialog()
}

#TESTA SE O COMPUTADOR ESTA CONECTADO NA REDE E RETORNA TRUE OU FALSE   
function testconnection($nomecomputador) {
        Test-Connection $nomecomputador -Quiet -Count 1            
    }

#RETORNA USUARIO LOGADO EM UM DETERMINADO COMPUTADOR
function UserName($computador){
    $CS = Gwmi Win32_ComputerSystem -Comp "$computador"
    return $CS.UserName
}

#ERRO - FALTA SELECIONAR UMA SEÇÃO
function ErroSecao{
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  

        #CONFIGURANDO JANELA
        $Form = New-Object System.Windows.Forms.Form
        $form.minimizebox = $false
        $form.maximizebox = $false
        $Form.Size = New-Object System.Drawing.Size(250, 100)
        $Form.Text = "ERRO" 
        $Form.StartPosition = "CenterScreen"

        #MENSAGEM DE ERRO
        $Titulo = New-Object System.Windows.Forms.Label
        $Titulo.Text = "Selecione uma seção"
        $Titulo.Size = New-Object System.Drawing.Size(120, 50)
        $Titulo.Location = New-Object System.Drawing.Size(30, 20)
        $Form.Controls.Add($Titulo)
        
        #FECHA A JANELA
        $Button_ok = New-Object System.Windows.Forms.Button 
        $Button_ok.Location = New-Object System.Drawing.Size(160, 10) 
        $Button_ok.Size = New-Object System.Drawing.Size(30, 30) 
        $Button_ok.text = "OK"
        $Button_ok.Add_Click( {$Form.Close()})
        $Form.Controls.Add($Button_ok)   
        
        #INICIA JANELA
        $result = $form.ShowDialog()

}

#ERRO - FALTA SELECIONAR UM usuario
function ErroUsuario{
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  

        #CONFIGURANDO JANELA
        $Form = New-Object System.Windows.Forms.Form
        $form.minimizebox = $false
        $form.maximizebox = $false
        $Form.Size = New-Object System.Drawing.Size(250, 100)
        $Form.Text = "ERRO" 
        $Form.StartPosition = "CenterScreen"

        #MENSAGEM DE ERRO
        $Titulo = New-Object System.Windows.Forms.Label
        $Titulo.Text = "Selecione um usuário"
        $Titulo.Size = New-Object System.Drawing.Size(120, 50)
        $Titulo.Location = New-Object System.Drawing.Size(30, 20)
        $Form.Controls.Add($Titulo)
        
        #FECHA A JANELA
        $Button_ok = New-Object System.Windows.Forms.Button 
        $Button_ok.Location = New-Object System.Drawing.Size(160, 10) 
        $Button_ok.Size = New-Object System.Drawing.Size(30, 30) 
        $Button_ok.text = "OK"
        $Button_ok.Add_Click( {$Form.Close()})
        $Form.Controls.Add($Button_ok)   
        
        #INICIA JANELA
        $result = $form.ShowDialog()

}

#ERRO - FALTA SELECIONAR UM COMPUTADOR
function ErroComputador {
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  

        #CONFIGURANDO JANELA
        $Form = New-Object System.Windows.Forms.Form
        $form.minimizebox = $false
        $form.maximizebox = $false
        $Form.Size = New-Object System.Drawing.Size(250, 100)
        $Form.Text = "ERRO"  
        $Form.StartPosition = "CenterScreen"

        #MENSAGEM DE ERRO
        $Titulo = New-Object System.Windows.Forms.Label
        $Titulo.Text = "Selecione um computador"
        $Titulo.Size = New-Object System.Drawing.Size(120, 50)
        $Titulo.Location = New-Object System.Drawing.Size(30, 20)
        $Form.Controls.Add($Titulo)
        
        #FECHA A JANELA
        $Button_ok = New-Object System.Windows.Forms.Button 
        $Button_ok.Location = New-Object System.Drawing.Size(160, 10) 
        $Button_ok.Size = New-Object System.Drawing.Size(30, 30) 
        $Button_ok.text = "OK"
        $Button_ok.Add_Click( {$Form.Close()})
        $Form.Controls.Add($Button_ok)   
    
        #INICIA A JANELA
        $result = $form.ShowDialog()

       

    }

#EXIBE PEQUENA MENSAGEM NA TELA
function ExibirMensagem($mensagem){
            [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
            [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  

            #CONFIGURAR JANELA
            $Form_selecione = New-Object System.Windows.Forms.Form
            $form_selecione.minimizebox = $false
            $form_selecione.maximizebox = $false
            $Form_selecione.Size = New-Object System.Drawing.Size(390, 100)
            $Form_selecione.Text = "NTI-JFAL"          
            $Form_selecione.StartPosition = "CenterScreen"

            #MENSAGEM EXIBIDA
            $Titulo = New-Object System.Windows.Forms.Label
            $Titulo.Text = $mensagem
            $Titulo.Size = New-Object System.Drawing.Size(250, 50)
            $Titulo.Location = New-Object System.Drawing.Size(20, 20)
            $Form_selecione.Controls.Add($Titulo)
            
            #FECHA A JANELA
            $Button_ok = New-Object System.Windows.Forms.Button 
            $Button_ok.Location = New-Object System.Drawing.Size(330,20) 
            $Button_ok.Size = New-Object System.Drawing.Size(30, 30) 
            $Button_ok.text = "OK"
            $Button_ok.Add_Click( {$Form_selecione.Close()})
            $Form_selecione.Controls.Add($Button_ok)
            
            #INICIA JANELA   
            $result = $form_selecione.ShowDialog()

}

function Users($computador){
    $c=Invoke-command -ComputerName NTI-035 -ScriptBlock {get-childItem -path c:\users}
    $c.name | ForEach-Object  {
        write-host "$_"
    }
}

#EXECUTA COMANDO GPResult EM UM DETERMINADO COMPUTADOR
function GPResult($computador){
     if(testconnection($computador)){
     
         #CONFIGURAR JANELA
         $Form = New-Object System.Windows.Forms.Form
         $form.minimizebox = $false
         $form.maximizebox = $false
         $Form.Size = New-Object System.Drawing.Size(235, 500)
         $Form.Text = "GPResult"          
         $Form.StartPosition = "CenterScreen"
     
         $Titulo = New-Object System.Windows.Forms.Label
         $Titulo.Text = "SELECIONE UM USUÁRIO"
         $Titulo.Location = New-Object System.Drawing.Size(10, 10)
         $Titulo.Size = New-Object System.Drawing.Size(250, 20)
         $Titulo.Font = New-Object System.Drawing.Font("IMPACT", 15, [System.Drawing.FontStyle]::Italic)
         $Form.Controls.Add($Titulo)
     
         $ListBox_usuarios = New-Object System.Windows.Forms.ListBox
         $ListBox_usuarios.Name = "Seções"
         $ListBox_usuarios.Width = 200
         $ListBox_usuarios.Height = 300
         $ListBox_usuarios.Location = New-Object System.Drawing.Size(10, 35)    
         $Form.Controls.Add($ListBox_usuarios)
          
         $outputBox = New-Object system.windows.Forms.TextBox
         $outputBox.Multiline = $true
         $outputBox.Size = New-Object System.Drawing.Size(200, 80)
         $outputBox.ReadOnly = $true
         $outputBox.ScrollBars = "Vertical"
         $outputBox.location = new-object system.drawing.point(10, 340)
         $outputBox.Font = "Lucida Sans,9"
         $Form.Controls.Add($outputBox)
     
         $Button_Cancelar = New-Object System.Windows.Forms.Button 
         $Button_Cancelar.Location = New-Object System.Drawing.Size(10, 430) 
         $Button_Cancelar.Size = New-Object System.Drawing.Size(100, 20)      
         $button_Cancelar.text = "Cancelar"
         $Button_Cancelar.Add_Click( {$Form.Close()}) 
         $Form.Controls.Add($Button_Cancelar)
     
     
         $c=Invoke-command -ComputerName $computador -ScriptBlock {get-childItem -path c:\users}
         $c.name | ForEach-Object  {
                           $ListBox_usuarios.Items.add($_)
                    }
         
         
         $Button_Confirmar = New-Object System.Windows.Forms.Button 
         $Button_Confirmar.Location = New-Object System.Drawing.Size(110, 430) 
         $Button_Confirmar.Size = New-Object System.Drawing.Size(100, 20)      
         $button_Confirmar.text = "Confirmar"
         $Button_Confirmar.Add_Click( {
            if($ListBox_usuarios.SelectedItem -ne $null){   
                $usuario=$ListBox_usuarios.SelectedItem
                $outputBox.Appendtext("INICIANDO...")
                $outputBox.Appendtext("`n")
                $destino= "GPResult_"+$computador+"_"+$usuario
                Get-GPResultantSetOfPolicy -Computer $computador -User $usuario -Path $env:userprofile\Desktop\$destino.html -ReportType html
                $outputBox.Appendtext("Tarefa concluida com sucesso")
                $outputBox.Appendtext("`n")
                $outputBox.Appendtext("Disponível em $env:userprofile\Desktop\$destino.html")         
            }
            else{
                ErroUsuario
            }
     
     
     
         }) 
         $Form.Controls.Add($Button_Confirmar)
     
        #INICIA JANELA   
        $result = $form.ShowDialog()
    }
    else{
          $mensagem = "Computador Desconectado"
          ExibirMensagem($mensagem)
            
         }

   

}

function PSDrive($computador){
    #CONFIGURAR JANELA
     $Form = New-Object System.Windows.Forms.Form
     $form.minimizebox = $false
     $form.maximizebox = $false
     $Form.Size = New-Object System.Drawing.Size(670, 550)
     $Form.Text = "PSDrive"          
     $Form.StartPosition = "CenterScreen"
     
     $Titulo = New-Object System.Windows.Forms.Label
     $Titulo.Text = "SELECIONE UM USUÁRIO"
     $Titulo.Location = New-Object System.Drawing.Size(10, 10)
     $Titulo.Size = New-Object System.Drawing.Size(250, 20)
     $Titulo.Font = New-Object System.Drawing.Font("IMPACT", 15, [System.Drawing.FontStyle]::Italic)
     $Form.Controls.Add($Titulo)
     
     $ListBox_usuarios = New-Object System.Windows.Forms.ListBox
     $ListBox_usuarios.Name = "Seções"
     $ListBox_usuarios.Width = 200
     $ListBox_usuarios.Height = 300
     $ListBox_usuarios.Location = New-Object System.Drawing.Size(10, 35)    
     $Form.Controls.Add($ListBox_usuarios)
          
     $outputBox = New-Object system.windows.Forms.TextBox
     $outputBox.Multiline = $true
     $outputBox.Size = New-Object System.Drawing.Size(400, 290)
     $outputBox.ReadOnly = $true
     $outputBox.ScrollBars = "Vertical"
     $outputBox.location = new-object system.drawing.point(230, 35)
     $outputBox.Font = "Lucida Sans,9"
     $Form.Controls.Add($outputBox)
     
     $Button_Cancelar = New-Object System.Windows.Forms.Button 
     $Button_Cancelar.Location = New-Object System.Drawing.Size(10, 330) 
     $Button_Cancelar.Size = New-Object System.Drawing.Size(100, 20)      
     $button_Cancelar.text = "Cancelar"
     $Button_Cancelar.Add_Click( {$Form.Close()}) 
     $Form.Controls.Add($Button_Cancelar)

     $ComboBox = New-Object System.Windows.Forms.ComboBox
     $COMBOBox.Text = "Unidade"
     $ComboBox.Width = 62
     $ComboBox.Height = 100
     $ComboBox.Location = New-Object System.Drawing.Size(230, 330)
     $Form.Controls.Add($ComboBox)

     $c=Invoke-command -ComputerName $computador -ScriptBlock {get-childItem -path c:\users}
     $c.name | ForEach-Object  {
                    $ListBox_usuarios.Items.add($_)
                }

     $Button_Confirmar = New-Object System.Windows.Forms.Button 
     $Button_Confirmar.Location = New-Object System.Drawing.Size(110, 330) 
     $Button_Confirmar.Size = New-Object System.Drawing.Size(100, 20)      
     $button_Confirmar.text = "Confirmar"
     $Button_Confirmar.Add_Click( {

        if ($ListBox_usuarios.SelectedIndices -ne $null){
                                 
            if(testconnection($computador)){   
                 $usuario=$ListBox_usuarios.SelectedItem
                 $outputBox.Appendtext("INICIANDO...")
                 $outputBox.Appendtext("`n")
                 $r =  Invoke-Command -ComputerName $computador -ScriptBlock {Get-CimInstance -Class Win32_NetworkConnection | format-list} | Out-String
                 $outputBox.Appendtext("$r")
                 
                           
            }
            else{
                 $outputBox.Appendtext("Computador Desconectado!")
            }
        }else {
                 ErroUsuario
        }
     
     }) 
     $Form.Controls.Add($Button_Confirmar)
     
    #INICIA JANELA   
    $result = $form.ShowDialog()

    #########
}
 
function Impressoras($computador){

                if(testconnection($computador)){   
                        Invoke-Command -ComputerName $computador -ScriptBlock {Get-CimInstance -Class Win32_Printer | format-list} | Out-String
                  
                           
                }else{
                    Write-Error ("Computador Desconectado!")
                }
}

function startmenu {
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  

        $res = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds

        $Form = New-Object System.Windows.Forms.Form
        $form.minimizebox = $true
        $form.maximizebox = $true
        $Form.Size = New-Object System.Drawing.Size(670, 550)
        $Form.Text = "NTI-JFAL v1.0"  
        $Form.StartPosition = "CenterScreen"


        $Font = New-Object System.Drawing.Font("IMPACT", 18, [System.Drawing.FontStyle]::Italic)

        $ListBox_secoes = New-Object System.Windows.Forms.ListBox
        $ListBox_secoes.Name = "Seções"
        $ListBox_secoes.Width = 200
        $ListBox_secoes.Height = 450
        $ListBox_secoes.Location = New-Object System.Drawing.Size(40, 20)    
        $Form.Controls.Add($ListBox_secoes)

        $ListBox_pcs = New-Object System.Windows.Forms.ListBox
        $ListBox_pcs.Name = "PCs"
        $ListBox_pcs.Width = 200
        $ListBox_pcs.Height = 450
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

        $Button_Impressoras = New-Object System.Windows.Forms.Button 
        $Button_Impressoras.Location = New-Object System.Drawing.Size(470, 410) 
        $Button_Impressoras.Size = New-Object System.Drawing.Size(150, 80) 
        $Button_Impressoras.font = $font
        $Button_Impressoras.forecolor = "Black"
        $Button_Impressoras.text = "Impressoras"
        $Button_Impressoras.Add_Click( {

        if ($ListBox_pcs.SelectedIndices -ne $null){
                    Impressoras($ListBox_secoes.SelectedItems)
                    #http://underpop.online.fr/w/windows-power-shell/trabalhando-com-impressoras-microsoft-windows-powershell.htm
                }
                else {
                    ErroComputador
                }    
                
                    
    
            }) 
        $Form.Controls.Add($Button_Impressoras) 

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


        $Button_PSDrive = New-Object System.Windows.Forms.Button 
        $Button_PSDrive.Location = New-Object System.Drawing.Size(470, 310) 
        $Button_PSDrive.Size = New-Object System.Drawing.Size(150, 80) 
        $Button_PSDrive.font = $font
        $Button_PSDrive.forecolor = "Black"
        $Button_PSDrive.text = "PSDrive"
        $Button_PSDrive.Add_Click( {
                #if ($ListBox_pcs.SelectedIndices -ne $null){
                    PSDrive($ListBox_pcs.SelectedItem)
                #}
                #else {
                 #   ErroComputador
                #}       
            }) 
        $Form.Controls.Add($Button_PSDrive)   


        $Button_Cancelar = New-Object System.Windows.Forms.Button 
        #$Button_Cancelar.Location = New-Object System.Drawing.Size(470, 410)
        $Button_Cancelar.Location = New-Object System.Drawing.Size(350, 470) 
        $Button_Cancelar.Size = New-Object System.Drawing.Size(100, 20) 
        #$button_Cancelar.font = $font
        $button_Cancelar.forecolor = "Black"
        $button_Cancelar.text = "Cancelar"
        $Button_Cancelar.Add_Click( {$Form.Close()}) 
        $Form.Controls.Add($Button_Cancelar)

        $Button_Computadores = New-Object System.Windows.Forms.Button 
        $Button_Computadores.Location = New-Object System.Drawing.Size(140, 470) 
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
        $Button_secoes.Location = New-Object System.Drawing.Size(40, 470) 
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


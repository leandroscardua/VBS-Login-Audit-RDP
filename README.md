# VBS-Login-Audit-RDP
Com esse script voce pode adcionar uma camada de controle no acesso as servidor Windows usando protocolo RDP. O script iram coletar informacoes como nome do usuario, nome de servidor de destino, nome do computador de origem e razao do acesso. O scrip ira criar um evento no Log de eventos do windows com as informacoes citadas anteriormente e se o servidor de email for configurado corretamente ira mandar esse informacao tambem.
O script vai ira executar a area de trabalho apos o usuario inserir a razao para o login no servidor ou estacao de trabalho


1 - Adcionar o script no registro: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\Userinit
    O caminho do script no Value data do registro, segue abaixo um exemplo, mas voce pode alterar de acordo com sua necessidade.
    C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -windowstyle hidden cscript c:\tmp\Login_Audit.vbs

2 - Configura a parte de servidor de SMTP, Porta e informacoes de envio de email,assunto e corpo do email

3 - Reinicie o Servidor ou Estacao de trabalho e teste o accesso.

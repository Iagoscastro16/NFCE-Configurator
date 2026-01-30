# =========================================================
# ConfigNFENew_Com_Interface.ps1
# =========================================================

# AVISO: Este script altera registros do Windows e de autoridades certificadoras. Use por sua conta e risco.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Função Principal ---
function Executar-Configuracao {
    param($OutputLabel)
    
    $OutputLabel.Text = "Status: Processando..."
    $OutputLabel.ForeColor = "Blue"
    
    # 1. Caminhos dos registros
    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
    $iePath = "HKCU:\Software\Microsoft\Internet Explorer\Download"
    $winTrustPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\WinTrust\Trust Providers\Software Publishing"

    # Garantir chaves
    foreach ($p in @($regPath, $iePath, $winTrustPath)) {
        if (!(Test-Path $p)) { New-Item -Path $p -Force | Out-Null }
    }

    # 2. Aplicando Registros
    Set-ItemProperty -Path $regPath -Name "SecureProtocols" -Value 2080
    Set-ItemProperty -Path $regPath -Name "DisableCachingOfSSLPages" -Value 1
    Set-ItemProperty -Path $regPath -Name "ManageSecurityCloud" -Value 0
    Set-ItemProperty -Path $regPath -Name "CertificateRevocation" -Value 0
    Set-ItemProperty -Path $iePath -Name "RunInvalidSignatures" -Value 1
    Set-ItemProperty -Path $iePath -Name "CheckExeSignatures" -Value "no"
    Set-ItemProperty -Path $winTrustPath -Name "State" -Value 146944

    # 3. Limpeza de Certificados
    $lojas = @("Cert:\LocalMachine\Root", "Cert:\CurrentUser\Root")
    foreach ($loja in $lojas) {
        $certs = Get-ChildItem $loja | Where-Object { $_.Subject -match "Autoridade Certificadora Raiz Brasileira v(1|2|5|10)" }
        if ($certs) {
            $certs | Remove-Item -ErrorAction SilentlyContinue
        }
    }

    $OutputLabel.Text = "Status: Concluído com Sucesso!"
    $OutputLabel.ForeColor = "DarkGreen"
    [System.Windows.Forms.MessageBox]::Show("As configurações de Internet e Certificados NFe foram aplicadas!", "Sucesso")
}

# --- Criação da Interface Gráfica ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Configurador NFe"
$form.Size = New-Object System.Drawing.Size(350,220)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Título
$label = New-Object System.Windows.Forms.Label
$label.Text = "Configuração Automática NFe"
$label.Location = New-Object System.Drawing.Point(20,20)
$label.Size = New-Object System.Drawing.Size(300,20)
$label.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($label)

# Label de Status
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Status: Aguardando comando..."
$statusLabel.Location = New-Object System.Drawing.Point(20,50)
$statusLabel.Size = New-Object System.Drawing.Size(300,20)
$form.Controls.Add($statusLabel)

# Botão ação
$btnExec = New-Object System.Windows.Forms.Button
$btnExec.Text = "APLICAR CONFIGURAÇÕES"
$btnExec.Location = New-Object System.Drawing.Point(50,90)
$btnExec.Size = New-Object System.Drawing.Size(230,45)
$btnExec.BackColor = "LightBlue"
$btnExec.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnExec.Add_Click({ Executar-Configuracao -OutputLabel $statusLabel })
$form.Controls.Add($btnExec)

# Nota de rodapé
$note = New-Object System.Windows.Forms.Label
$note.Text = "Execute como Administrador para melhores resultados."
$note.Location = New-Object System.Drawing.Point(20,150)
$note.Size = New-Object System.Drawing.Size(300,20)
$note.Font = New-Object System.Drawing.Font("Arial", 8)
$form.Controls.Add($note)

# Exibição da janela
$form.ShowDialog() | Out-Null
# =========================================================
# ConfigNFCeNew Com Interface
# =========================================================

# AVISO: Este script altera registros do Windows e de autoridades certificadoras. Use por sua conta e risco.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Função de Log ---
function Write-Log {
    param(
        [string]$Mensagem,
        [ValidateSet("OK","ERRO","INFO","AVISO")][string]$Tipo = "INFO"
    )

    $logPath = Join-Path $PSScriptRoot "NFCe_log.txt"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $linha = "[$timestamp] [$Tipo] $Mensagem"

    Add-Content -Path $logPath -Value $linha -Encoding UTF8
}

# --- Função Principal ---
function Executar-Configuracao {
    param($OutputLabel)

    $OutputLabel.Text = "Status: Processando..."
    $OutputLabel.ForeColor = "Blue"

    $erros = 0

    Write-Log "========================================" "INFO"
    Write-Log "Iniciando configuração NFCe" "INFO"
    Write-Log "Usuário: $env:USERNAME | Máquina: $env:COMPUTERNAME" "INFO"

    # 1. Caminhos dos registros
    $regPath      = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
    $iePath       = "HKCU:\Software\Microsoft\Internet Explorer\Download"
    $winTrustPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\WinTrust\Trust Providers\Software Publishing"

    # Garantir que as chaves existem
    foreach ($p in @($regPath, $iePath, $winTrustPath)) {
        try {
            if (!(Test-Path $p)) {
                New-Item -Path $p -Force | Out-Null
                Write-Log "Chave de registro criada: $p" "OK"
            } else {
                Write-Log "Chave de registro já existe: $p" "INFO"
            }
        } catch {
            Write-Log "Falha ao verificar/criar chave: $p — $_" "ERRO"
            $erros++
        }
    }

    # 2. Aplicando Registros — cada um com try/catch individual
    $registros = @(
        @{ Path=$regPath;      Name="SecureProtocols";        Value=2080  },
        @{ Path=$regPath;      Name="DisableCachingOfSSLPages"; Value=1   },
        @{ Path=$regPath;      Name="ManageSecurityCloud";     Value=0    },
        @{ Path=$regPath;      Name="CertificateRevocation";   Value=0    },
        @{ Path=$iePath;       Name="RunInvalidSignatures";    Value=1    },
        @{ Path=$iePath;       Name="CheckExeSignatures";      Value="no" },
        @{ Path=$winTrustPath; Name="State";                   Value=146944 }
    )

    foreach ($reg in $registros) {
        try {
            Set-ItemProperty -Path $reg.Path -Name $reg.Name -Value $reg.Value -ErrorAction Stop
            Write-Log "Registro aplicado: $($reg.Name) = $($reg.Value)" "OK"
        } catch {
            Write-Log "Falha ao aplicar registro: $($reg.Name) — $_" "ERRO"
            $erros++
        }
    }

    # 3. Limpeza de Certificados
    Write-Log "Iniciando limpeza de certificados raiz brasileiros..." "INFO"

    $lojas = @("Cert:\LocalMachine\Root", "Cert:\CurrentUser\Root")
    foreach ($loja in $lojas) {
        try {
            $certs = Get-ChildItem $loja -ErrorAction Stop |
                     Where-Object { $_.Subject -match "Autoridade Certificadora Raiz Brasileira v(1|2|5|10)" }

            if ($certs) {
                foreach ($cert in $certs) {
                    try {
                        $cert | Remove-Item -ErrorAction Stop
                        Write-Log "Certificado removido [$loja]: $($cert.Subject)" "OK"
                    } catch {
                        Write-Log "Falha ao remover certificado [$loja]: $($cert.Subject) — $_" "ERRO"
                        $erros++
                    }
                }
            } else {
                Write-Log "Nenhum certificado alvo encontrado em: $loja" "AVISO"
            }
        } catch {
            Write-Log "Falha ao acessar loja de certificados: $loja — $_" "ERRO"
            $erros++
        }
    }

    # 4. Resultado final
    if ($erros -eq 0) {
        Write-Log "Configuração concluída sem erros." "OK"
        $OutputLabel.Text = "Status: Concluído com Sucesso!"
        $OutputLabel.ForeColor = "DarkGreen"
        [System.Windows.Forms.MessageBox]::Show(
            "Configurações aplicadas com sucesso!`nLog salvo em: NFCe_log.txt",
            "Sucesso", "OK", "Information"
        )
    } else {
        Write-Log "Configuração concluída com $erros erro(s). Verifique o log." "AVISO"
        $OutputLabel.Text = "Status: Concluído com $erros erro(s). Veja o log."
        $OutputLabel.ForeColor = "DarkRed"
        [System.Windows.Forms.MessageBox]::Show(
            "Configuração finalizada, mas $erros item(ns) falharam.`nConsulte o arquivo NFCe_log.txt para detalhes.",
            "Atenção", "OK", "Warning"
        )
    }
}

# --- Criação da Interface Gráfica ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Configurador NFCe"
$form.Size = New-Object System.Drawing.Size(350,220)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Título
$label = New-Object System.Windows.Forms.Label
$label.Text = "Configuração Automática NFCe"
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
INSTRUÇÕES DE USO - CONFIGURADOR NFCe (Versão GUI)

## Como Executar
Para facilitar, criei um disparador automático. 
1. Baixe o conteúdo do repositório.
2. Localize o arquivo **`rodar_agora.bat`**.
3. Clique com o botão direito e selecione **"Executar como Administrador"**.

Este script foi desenvolvido para automatizar a configuração técnica de emissão de Notas Fiscais Eletrônicas, garantindo que o Windows esteja nos padrões exigidos pelos órgãos certificadores e SEFAZ, nele eu configuro as opções de internet avançadas, e removo as autoridades certificadores brasileiras de todas as versões usadas.


 Testado em: Windows 10 e Windows 11. 
 Compatibilidade Provável: Windows 8 e 8.1 (pendente de testes formais). 
 Não compatível: Windows 7 e versões anteriores.

🟦 1. rodar_agoranew.bat (O Disparador)
Este arquivo funciona como a chave de ignição do sistema. Ele foi configurado com quatro pilares de segurança e performance:

@echo off: Garante que o console não fique "sujo" com linhas de comando confusas, focando apenas na execução.

-NoProfile: Força o PowerShell a carregar um ambiente "limpo" (padrão de fábrica). Isso evita conflitos com configurações personalizadas ou scripts de terceiros que o usuário possa ter na máquina.

-ExecutionPolicy Bypass: Contorna a trava de segurança padrão do Windows que impede a execução de scripts .ps1. O Bypass abre uma exceção temporária apenas para este script, garantindo que o configurador rode sem interrupções.

%~dp0: Este código inteligente diz ao Windows: "Não importa em qual pasta o usuário salvou isso, procure o script de configuração exatamente no mesmo lugar onde este .bat está".

🟦 2. ConfigNFENew.ps1 (O Motor e a Interface)
Diferente das versões anteriores de terminal, este script agora possui uma Interface Gráfica (GUI) baseada em Windows Forms, trazendo uma experiência mais intuitiva e segura.

O que ele faz tecnicamente:

Interface Retrô: Cria uma janela nativa que evita que o usuário precise interagir com linhas de comando, reduzindo erros de operação.

Ajuste de Registros (Internet Settings): * Configura o SecureProtocols para 2080 (SSL 3.0 + TLS 1.2), o padrão exigido pela NFCe.

Desativa o cache de páginas SSL e o SmartScreen para evitar bloqueios de download de XML.

Permite a execução de softwares com assinaturas que o Windows julga "inválidas", essencial para certos componentes de certificados.

Limpeza Automática de Raízes: Localiza e remove silenciosamente os certificados da "Autoridade Certificadora Raiz Brasileira" (v1, v2, v5 e v10) tanto no repositório da Máquina Local quanto do Usuário Atual. Isso resolve problemas comuns de cadeias de certificados corrompidas ou duplicadas.

WinTrust State: Ajusta o estado de confiança do Windows para ignorar avisos burocráticos que travam a comunicação com os servidores da SEFAZ.

⚠️ AVISO IMPORTANTE
Para que o script consiga remover certificados da Raiz do Sistema e alterar chaves de registro protegidas, o arquivo rodar_agoranew.bat deve ser executado como Administrador.



## Isenção de Responsabilidade (Disclaimer)

Este script foi desenvolvido para fins de automação e facilitação técnica. Ao utilizá-lo, você está ciente de que:
* O script realiza alterações em chaves do **Registro do Windows (Regedit)**.
* O script remove certificados específicos da loja de raízes do sistema.
* **Uso por sua conta e risco:** Não me responsabilizo por eventuais instabilidades ou comportamentos inesperados no sistema operacional. 
* **Recomendação:** É altamente aconselhável criar um **Ponto de Restauração** do Windows antes de aplicar as configurações ou ter tirar um print das opções avançadas de internet e acessar o s.

> **Aviso:** Este script foi desenvolvido para o contexto tributário e técnico do Brasil (ICP-Brasil / SEFAZ).

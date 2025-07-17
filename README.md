# Remove Office Script

Este script PowerShell automatiza a remoção completa de versões não autorizadas (crackeadas) do Microsoft Office em máquinas que possuem o Microsoft 365 instalado corretamente.

---

## Objetivo

- Garantir que o Office crackeado seja completamente removido do cliente.
- Preparar o ambiente para uso legítimo do Office 365.
- Facilitar a execução remota via RMM ou outro sistema de gerenciamento.

---

## Como funciona

1. Baixa a ferramenta oficial **Office Deployment Tool (ODT)** da Microsoft.
2. Cria um arquivo de configuração XML para remover todas as versões do Office instaladas.
3. Executa a remoção silenciosa usando a ODT.
4. (Opcional) Pode ser executado remotamente puxando o script direto do GitHub.

---

## Como usar

### Pré-requisitos

- Windows 10 ou superior.
- PowerShell com permissões de administrador.
- Conexão com a internet para baixar a ODT.

### Execução manual

1. Baixe o script `remove_office.ps1`.
2. Abra o PowerShell como administrador.
3. Execute o script:

```powershell
.\remove_office.ps1

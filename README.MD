# 🛡️ Stable Excel Auto-Updater

Uma ferramenta de automação resiliente para atualização de dados em planilhas Excel, projetada para ambientes corporativos com latência de rede ou sistemas de arquivos legados.

![Python](https://img.shields.io/badge/python-3.9%2B-blue)
![Status](https://img.shields.io/badge/status-production-green)

## 🎯 Objetivo

Automatizar o fluxo de *Refresh* (atualização de conexões de dados) do Excel garantindo a integridade do arquivo através de uma **Estratégia de Sincronização Híbrida**.

## ⚙️ Arquitetura de Sincronização

Diferente de scripts comuns, esta solução não confia apenas nos callbacks do Excel. Ela implementa **Buffers de Estabilidade** (Time Delays) estratégicos para mitigar falhas de I/O (Entrada/Saída):

1.  **Buffer de Carregamento (5s):** Garante que arquivos pesados ou hospedados em rede sejam totalmente carregados na RAM antes de qualquer manipulação.
2.  **Sincronização de Consulta:** Utiliza `CalculateUntilAsyncQueriesDone()` para aguardar o término lógico das queries (Power Query/SQL).
3.  **Buffer de Persistência (5s):** Garante que o Sistema Operacional tenha tempo de escrever os dados no disco físico após o comando de salvar, prevenindo corrupção de arquivo ao fechar a aplicação imediatamente.
4.  **Cooldown Final (5s):** Aguarda o encerramento total do processo `EXCEL.EXE` no Windows para evitar conflitos em execuções sequenciais.

## 🛠️ Funcionalidades

* **Tratamento de Erros:** Uso de blocos `try/finally` para garantir que o Excel seja fechado mesmo se o script falhar, evitando consumo desnecessário de memória.
* **Logging:** Sistema de logs detalhados com *timestamps* para auditoria de execução.
* **Type Safety:** Validação de caminhos de arquivo antes da execução.

## 🚀 Como Usar

1. Instale a dependência de comunicação com o Windows:
   ```bash
   pip install pywin32

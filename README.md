# Outlook Structured Operations Audit Automation

> Bulk dispatch and follow-up automation for structured operations audit using Outlook and Excel.

---

## 📌 Visão Geral

Este projeto automatiza o envio e o acompanhamento de e-mails de **auditoria de operações estruturadas**, utilizando o Microsoft Outlook para comunicação e arquivos Excel como base de dados e trilha de auditoria.

Ele foi desenvolvido para suportar rotinas corporativas de auditoria nas quais diversos profissionais precisam, periodicamente, **confirmar ou validar operações**, garantindo:

- Rastreabilidade completa dos envios
- Histórico centralizado
- Controle de respostas
- Ciclos de cobrança automatizados

> ⚠️ Este é um projeto **sanitizado e adaptado para portfólio**, baseado em uma automação real utilizada em ambiente corporativo. Nenhuma informação sensível, dado real de cliente ou regra proprietária está incluída neste repositório.

---

## 🎯 Objetivo

Em muitos ambientes corporativos, processos de auditoria dependem de:

- Envio manual de e-mails
- Controle manual de quem respondeu e quem não respondeu
- Reenvio manual de cobranças
- Atualização manual de planilhas de controle

Este projeto resolve esse problema fornecendo:

- Envio em massa de e-mails via Outlook
- Geração de token único por registro auditado
- Registro automático de todos os envios em planilha de histórico
- Processo automatizado de follow-up e cobrança

---

## 🚀 Funcionalidades

- Envio em massa de e-mails de auditoria via Microsoft Outlook
- Geração de token único por registro para rastreabilidade
- Registro de todos os envios em arquivo Excel de histórico
- Automação de follow-up:
  - Busca respostas no Outlook
  - Marca registros como respondidos
  - Reenvia solicitações quando não há resposta
- Proteção contra uso simultâneo de arquivos Excel (evita corrupção de arquivos)
- Configuração centralizada via arquivo JSON
- Opção de:
  - Apenas exibir os e-mails antes do envio
  - Ou enviar automaticamente

---

## 🧱 Estrutura do Projeto

```text
outlook-structured-operations-audit-automation/
  main.py
  config.example.json
  requirements.txt
  README.md
  src/
    outlook_audit/
      config.py
      dispatch.py
      followup.py
      outlook_client.py
      history_store.py
      file_lock.py
      logging_utils.py

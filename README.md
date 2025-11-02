# ⚙️ Augmentir Access Automator (Python)

Automação desenvolvida para otimizar o processo de liberação de acessos na plataforma **[Augmentir](https://app.augmentir.com)**
Este projeto nasceu de uma necessidade real no estágio de **Sistemas de Manufatura**, onde o processo manual de verificação e criação de acessos levava até **45 minutos**.  
Com esta automação em **Python + Selenium**, o mesmo fluxo é executado em **menos de 10 minutos**, com precisão e logs automáticos.

---

## 🚀 O que o script faz

- Lê automaticamente as planilhas de solicitações de acesso da pasta `INPUT`
- Processa os dados e atualiza a **base mestre (Excel)** local
- Abre o **Augmentir** via Selenium WebDriver
- Verifica se cada funcionário já possui acesso
- Exibe resultados em tempo real (acesso existente / não encontrado)
- Move automaticamente os arquivos processados para a pasta `ARQUIVADOS`

---

## 🧩 Estrutura de Pastas

```bash
📁 Projeto/
├── INPUT/                # Planilhas recebidas por e-mail
├── ARQUIVADOS/           # Planilhas já processadas
├── Base de Dados da Solicitação de Acesso Sistemas de Manufatura.xlsx
├── augmentir_automator.py
└── README.md
```

---

## ⚙️ Tecnologias Utilizadas

- Python 3.11+
- Selenium WebDriver
- Pandas
- OpenPyXL
- ChromeDriver

---

## 🧠 Lógica Principal

- Processamento de planilhas
Lê todos os arquivos .xlsx em INPUT, extrai os nomes dos usuários e anexa os dados à base principal.
- Verificação no Augmentir
Usa Selenium para navegar até https://app.augmentir.com/#/configure/users, autenticar e verificar se cada nome já possui acesso.
- Resultados em tempo real
O script mostra no terminal:
✅ “Já possui acesso”
🚫 “Não possui acesso”
- Finalização
Ao final, todos os arquivos processados são movidos para ARQUIVADOS, e o robô exibe um resumo da execução.

---

## 🧭 Roadmap de Melhorias

- Geração automática de log (.txt) com data/hora
- Integração com API interna para atualização direta
- Relatório em Excel com nomes sem acesso
- Modo “headless” (execução invisível do navegador)
- Dashboard simples (Streamlit ou Flask)
- Parametrização de e-mail corporativo e planilhas via .env

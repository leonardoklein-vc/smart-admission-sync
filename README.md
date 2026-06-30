# 🏥 Smart Admission Sync (Gestor de Internação)

![Google Apps Script](https://img.shields.io/badge/Google_Apps_Script-4285F4?style=for-the-badge&logo=google&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
![Bootstrap 5](https://img.shields.io/badge/Bootstrap_5-7952B3?style=for-the-badge&logo=bootstrap&logoColor=white)

Uma Single Page Application (SPA) injetada diretamente no Google Sheets para gestão de leitos e conciliação automática de agendas hospitalares. 

---

## 🚨 O Problema: O Custo do "Ctrl+C, Ctrl+V"

Na rotina administrativa hospitalar, a equipe de internação gastava de **2 a 3 horas diárias** apenas para manter a sua agenda sincronizada com a agenda do Bloco Cirúrgico.
O processo era inteiramente manual e visual:
* Copiar dados de uma planilha para a outra.
* Procurar linha por linha quem foi adicionado (pacientes novos).
* Encontrar quem teve o horário da cirurgia alterado.
* Identificar pacientes que foram cancelados ou removidos da lista original.

Além do tempo exaustivo, o cansaço visual gerava um alto risco de falha humana em dados críticos.

*(Abaixo: A visão anterior — uma planilha crua e dependente de formatação manual)*
<br>
`image_62f1b7.png`

---

## 💡 A Solução: Motor de Sincronização (Diff Checker)

Desenvolvi um painel interativo que atua como um sistema de gestão completo por cima da base de dados do Google Sheets, equipado com um algoritmo de **Conciliação de Dados**.

O tempo de trabalho caiu de **horas para segundos**.

### ✨ Funcionalidades Principais

#### 1. Coleta e Atualização Inteligente (Sync Engine)
Ao invés de copiar e colar, o usuário clica em "Coletar Informações". O script vai até a planilha do Bloco Cirúrgico, varre os dados, processa regras de negócio estritas (identificando blocos como CCA, HMD) e compara com a agenda de internação atual, gerando um *Diff* visual em 3 categorias:
* 🟢 **NOVOS:** Pacientes que estão na cirurgia, mas não na internação (Sugere: *Inserir*).
* 🟣 **ALTERADOS:** Pacientes que já existem em ambas, mas com divergências de horário, convênio ou equipe (Sugere: *Atualizar com checkbox granular do que aceitar*).
* 🔴 **REMOVIDOS:** Pacientes que sumiram da lista cirúrgica (Sugere: *Excluir*).

*(Abaixo: O Motor de Sincronização em ação)*
<br>
`image_62f57d.png`

#### 2. Dashboard de Gestão (UI/UX)
A interface transforma a planilha em um software real.
* **Agenda Visual:** Tabela interativa com dropdowns de status coloridos.
* **Gestão de Leitos:** Visão clara de leitos indisponíveis e mapa de disponibilidades por andar.
* **Indicadores em Tempo Real:** Cards dinâmicos mostrando volumes de transferências e ocupação do dia.

*(Abaixo: A nova interface da aplicação)*
<br>
`image_62f4c1.png`

---

## ⚙️ Destaques da Arquitetura
* **Lógica de Slots (Vagas):** O backend possui um algoritmo que mapeia linhas vazias na planilha legada antes de inserir novos dados, evitando sobreposição e garantindo formatação automática do fundo e fonte.
* **Client-Side/Server-Side Handshake:** O uso do `google.script.run` para buscar o *Diff* no backend e renderizá-lo no frontend, permitindo que o usuário aprove as mudanças antes de efetivar o `commit` no banco de dados.
* **Color Mapping Dinâmico:** Um dicionário hexadecimal (`COLOR_MAP`) garante que a taxonomia visual do hospital seja respeitada automaticamente pelo sistema.

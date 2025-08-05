<div align="center">
  <h1>Assist RAT Downloader</h1>
  <p>
    Ferramenta de automação para download e organização de Relatórios de Atendimento Técnico (RATs) do portal Positivo.
  </p>

  <img src="https://img.shields.io/badge/Python-3.7%2B-blue?logo=python" alt="Python Version">
  <img src="https://img.shields.io/badge/License-MIT-green.svg" alt="License MIT">
  <img src="https://img.shields.io/badge/Status-Ativo-brightgreen" alt="Project Status">

</div>

---

> **Nota:** Esta ferramenta foi desenvolvida como uma prova de conceito para otimização de processos e para fins educacionais. Contém um aviso de segurança importante sobre a plataforma alvo. Leia a seção [**⚠️ Aviso de Segurança**](#-aviso-de-segurança) para mais detalhes.

---

### Tabela de Conteúdos
* [Sobre o Projeto](#sobre-o-projeto)
* [Principais Features](#principais-features)
* [Screenshots](#screenshots)
* [Construído Com](#construído-com)
* [Começando](#começando)
  * [Pré-requisitos](#pré-requisitos)
  * [Instalação](#instalação)
* [Como Usar](#como-usar)
* [⚠️ Aviso de Segurança](#-aviso-de-segurança)
* [Roadmap](#roadmap)
* [Como Contribuir](#como-contribuir)
* [Licença](#licença)
* [Autores](#autores)

---

## Sobre o Projeto

O processo de baixar relatórios de atendimento um por um é repetitivo, demorado e sujeito a erros. O **Assist RAT Downloader** foi criado para resolver esse problema, automatizando completamente o download, a organização e a compilação de centenas de relatórios com apenas alguns cliques.

A ferramenta lê uma planilha simples onde os chamados são organizados por técnico, baixa os PDFs correspondentes de forma paralela e os organiza em pastas, economizando horas de trabalho manual.

## Principais Features

* ✨ **Interface Gráfica Intuitiva:** Fácil de usar, sem necessidade de conhecimento técnico.
* ⚡️ **Downloads em Paralelo:** Utiliza múltiplos workers para baixar vários arquivos simultaneamente, maximizando a velocidade.
* 📂 **Organização Automática:** Cria uma pasta para cada técnico e salva os relatórios correspondentes.
* 📎 **Compilador de PDFs:** Une todos os PDFs de um técnico em um único arquivo, facilitando o compartilhamento e a análise.
* 🔄 **Modo Incremental e de Limpeza:** Escolha entre baixar apenas os arquivos novos ou apagar tudo e começar do zero.
* ✅ **Validação de Dados:** Analisa a planilha em busca de erros comuns (como chamados duplicados) antes de iniciar o processo.

## Screenshots

<div align="center">
  <img src="URL_DA_SUA_IMAGEM_OU_GIF_AQUI" alt="Demonstração do App" width="600"/>
</div>


## Construído Com

* [Python](https://www.python.org/)
* [Tkinter](https://docs.python.org/3/library/tkinter.html) & [ttkthemes](https://github.com/RedFantom/ttkthemes)
* [Requests](https://requests.readthedocs.io/en/latest/)
* [Pandas](https://pandas.pydata.org/)
* [pypdf](https://pypdf.readthedocs.io/en/stable/)

---

## Começando

Siga os passos abaixo para ter uma cópia do projeto rodando na sua máquina local.

### Pré-requisitos

* **Python 3.7 ou superior.**
* **Git** (opcional, para clonar o repositório).

### Instalação

1.  **Clone o repositório (ou baixe os arquivos):**
    ```sh
    git clone [https://github.com/SEU_USUARIO/SEU_REPOSITORIO.git](https://github.com/SEU_USUARIO/SEU_REPOSITORIO.git)
    cd SEU_REPOSITORIO
    ```

2.  **Crie um ambiente virtual (recomendado):**
    ```sh
    python -m venv venv
    source venv/bin/activate  # No Windows: venv\Scripts\activate
    ```

3.  **Instale as dependências:**
    O script tentará instalar as dependências automaticamente. Caso falhe, instale-as a partir do arquivo `requirements.txt`:
    ```sh
    pip install -r requirements.txt
    ```

---

## Como Usar

1.  **Prepare a Planilha `chamados.xlsx`:**
    Este é o passo mais importante. Crie ou edite o arquivo `chamados.xlsx` na mesma pasta do programa. A estrutura deve ser:
    * **Linha 1:** O nome de cada técnico em uma coluna.
    * **Linhas Seguintes:** Os números dos chamados (RATs) abaixo do técnico correspondente.

    **Exemplo de Estrutura:**
    | **Fulano da Silva** | **Ciclano de Souza** | **Beltrano de Oliveira** |
    |-------------------|--------------------|------------------------|
    | 500123456         | 500789012          | 500654321              |
    | 500234567         | 500890123          | 500765432              |
    | ...               | ...                | ...                    |

2.  **Execute o Programa:**
    ```sh
    python main.py
    ```

3.  **Use a Interface:**
    * Clique em **"Selecionar Pasta de Destino"** e escolha onde os arquivos serão salvos.
    * Marque a opção **"Apagar arquivos anteriores"** se desejar o modo de limpeza total.
    * Clique em **"Iniciar"**.
    * Acompanhe o progresso na caixa de "Log de Atividades".

---

## ⚠️ Aviso de Segurança

Durante o desenvolvimento desta ferramenta, foi identificada uma vulnerabilidade de **Controle de Acesso Quebrado (Broken Access Control)** na plataforma `assist.positivotecnologia.com.br`.

### A Falha

> A aplicação não exige autenticação (login) para permitir o download dos Relatórios de Atendimento Técnico (RATs). O acesso ao endpoint que gera os PDFs (`.../gerarRatPdf.php`) requer apenas um *cookie de sessão* válido, que é obtido simplesmente ao visitar a página inicial do portal, sem necessidade de credenciais.

### Implicação

> Esta falha permite que qualquer pessoa com posse de um ID de chamado (`os_id`) válido possa construir a URL de download e baixar o relatório correspondente. Isso representa um risco de vazamento de informações, pois dados potencialmente sensíveis dos relatórios podem ser acessados por terceiros não autorizados.

**Aviso de Divulgação Responsável:**
*Esta informação é compartilhada estritamente para fins educacionais e de conscientização sobre segurança da informação. A ferramenta foi projetada para otimizar processos legítimos. O uso indevido desta informação ou da ferramenta é de inteira responsabilidade do usuário. Recomenda-se fortemente que os administradores do sistema alvo corrijam esta vulnerabilidade, aplicando a verificação de sessão autenticada para o acesso a recursos restritos.*

---

## Roadmap

* [ ] Adicionar um painel de configuração dentro da GUI.
* [ ] Suporte para importação de chamados a partir de arquivos `.csv`.
* [ ] Empacotar a aplicação em um executável único para Windows, macOS e Linux.
* [ ] Melhorar o relatório final, com estatísticas mais detalhadas.

Veja as [issues abertas](https://github.com/SEU_USUARIO/SEU_REPOSITORIO/issues) para uma lista completa de funcionalidades propostas (e bugs conhecidos).

## Como Contribuir

Contribuições são o que tornam a comunidade de código aberto um lugar incrível para aprender, inspirar e criar. Qualquer contribuição que você fizer será **muito apreciada**.

1.  Faça um **Fork** do Projeto.
2.  Crie sua **Feature Branch** (`git checkout -b feature/AmazingFeature`).
3.  Faça um **Commit** de suas mudanças (`git commit -m 'Add some AmazingFeature'`).
4.  Faça um **Push** para a Branch (`git push origin feature/AmazingFeature`).
5.  Abra um **Pull Request**.

## Licença

Distribuído sob a Licença MIT. Veja `LICENSE.txt` para mais informações.

## Autores

* **Sojo & Cabelo** - *Desenvolvimento*
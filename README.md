<div align="center">

<img src="assets/E-Fal.png" alt="E-Fal Logo" width="160"/>

# E-Fal

**Emissão Automatizada de Certidões de Falência — TJRJ**

[![Python](https://img.shields.io/badge/Python-3.6%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![Selenium](https://img.shields.io/badge/Selenium-WebDriver-43B02A?style=for-the-badge&logo=selenium&logoColor=white)](https://www.selenium.dev/)
[![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white)](https://www.microsoft.com/windows)
[![License](https://img.shields.io/badge/License-MIT-yellow?style=for-the-badge)](LICENSE)
[![Chrome](https://img.shields.io/badge/Google%20Chrome-Required-4285F4?style=for-the-badge&logo=googlechrome&logoColor=white)](https://www.google.com/chrome/)

----

Ferramenta de automação para emissão de certidões de falência no portal do
**Tribunal de Justiça do Estado do Rio de Janeiro (TJRJ)**, eliminando a interação manual
com a interface web e acelerando processos jurídicos em lote.

[Funcionalidades](#-funcionalidades) &bull;
[Instalação](#-instalação) &bull;
[Como Usar](#-como-usar) &bull;
[Configuração](#%EF%B8%8F-configuração) &bull;
[Contribuição](#-contribuição)

</div>

---

## Sobre o Projeto

O **E-Fal** foi desenvolvido para automatizar o fluxo de emissão de certidões de falência junto ao TJRJ. A aplicação lê uma planilha Excel contendo dados de pedidos, preenche automaticamente os formulários do portal judicial, e realiza o download dos PDFs das certidões — tudo com acompanhamento em tempo real por meio de uma interface gráfica intuitiva.

### Tecnologias Utilizadas

| Tecnologia | Função |
|---|---|
| **Python 3.6+** | Linguagem principal |
| **Tkinter** | Interface gráfica (GUI) |
| **Selenium** | Automação do navegador |
| **openpyxl** | Leitura e validação de arquivos Excel |
| **WebDriver Manager** | Gerenciamento automático do ChromeDriver |
| **psutil** | Controle de processos do Chrome |
| **Pillow** | Manipulação de imagens na GUI |
| **PyAutoGUI** | Automação de teclado/mouse |

---

## Funcionalidades

- **Validação de Excel** — Verifica se o arquivo fornecido possui as colunas esperadas (`Código`, `Cliente`, `CND`, `Pedido`) e exibe os dados na interface para conferência.
- **Processamento em Lote** — Preenche formulários no portal do TJRJ, consulta certidões e realiza o download dos PDFs automaticamente para múltiplos pedidos.
- **Download Gerenciado** — Salva os arquivos na pasta de destino com nomenclatura padronizada baseada no código do pedido e na data de vencimento.
- **Interface Gráfica** — Painel intuitivo com seleção de arquivos, barra de progresso, tabela de dados e log de execução em tempo real.
- **Log de Execução** — Registra todas as ações em arquivo de log com timestamp para rastreabilidade e depuração.
- **Controle de Sessão** — Utiliza perfil persistente do Chrome, mantendo sessões autenticadas entre execuções.

---

## Pré-requisitos

| Requisito | Detalhes |
|---|---|
| Sistema Operacional | Windows 10/11 |
| Python | Versão 3.6 ou superior |
| Navegador | Google Chrome (versão atualizada) |
| Conexão | Internet estável durante o processamento |

---

## Instalação

### 1. Clone o repositório

```bash
git clone https://github.com/seu-usuario/E-Fal.git
cd E-Fal
```

### 2. Instale as dependências

```bash
pip install openpyxl selenium webdriver-manager psutil Pillow pyautogui
```

### 3. Verifique o Chrome

Certifique-se de que o Google Chrome está instalado e atualizado. O `webdriver-manager` cuidará automaticamente do download do ChromeDriver compatível.

---

## Como Usar

### 1. Execute a aplicação

```bash
python e-fal.py
```

### 2. Na interface gráfica

| Passo | Ação |
|:---:|---|
| **1** | Clique em **Procurar** para selecionar o arquivo Excel (`.xlsx`) |
| **2** | Defina a **pasta de destino** onde os PDFs serão salvos |
| **3** | Clique em **Validar Excel** para verificar e visualizar os dados |
| **4** | Clique em **Iniciar Processamento** para começar a automação |
| **5** | Acompanhe o progresso pelo log e pela barra de progresso |

> **Dica:** Na primeira execução, pode ser necessário fazer login manualmente no portal do TJRJ. O perfil do Chrome persistirá a sessão para as próximas execuções.

---

## Formato do Excel

O arquivo Excel deve conter as seguintes colunas na primeira linha:

| Coluna | Descrição | Obrigatório |
|---|---|:---:|
| `Código` | Identificador único do pedido | Sim |
| `Cliente` | Nome do cliente | Sim |
| `CND` | Tipo da certidão | Não |
| `Pedido` | Número do pedido no TJRJ | Sim |

**Exemplo:**

| Código | Cliente | CND | Pedido |
|---|---|---|---|
| 12345 | João Silva | CND FALENCIA | 2023-0012345 |
| 67890 | Maria Oliveira | CND FALENCIA | 2023-0067890 |

---

## Configuração

| Parâmetro | Valor Padrão | Descrição |
|---|---|---|
| **URL do TJRJ** | `https://www3.tjrj.jus.br/CJE/certidao/judicial/visualizar?modelo=visualizar` | Endereço do portal de certidões (editável na interface) |
| **Pasta de Destino** | `C:\Users\VM001\Documents\CNDs\FALENCIA` | Local de salvamento dos PDFs |
| **Perfil do Chrome** | `C:\PerfisChrome\automacao\Profile 1` | Perfil para persistência de sessão |
| **Pasta de Logs** | `./TJRJ_Logs/` | Diretório de logs (criado automaticamente) |

---

## Estrutura do Projeto

```
E-Fal/
├── assets/
│   ├── E-Fal.ico          # Ícone da aplicação
│   └── E-Fal.png          # Logo exibido na interface
├── Old_v/
│   └── testeInterface.py   # Versão anterior (CustomTkinter)
├── e-fal.py                # Aplicação principal
├── .gitignore
├── LICENSE
└── README.md
```

---

## Limitações

- Projetado exclusivamente para o portal do **TJRJ** — alterações na estrutura do site podem exigir ajustes nos seletores XPath.
- Requer **Google Chrome** instalado no sistema.
- Funciona apenas em **Windows** devido às dependências de caminhos e automação de teclado.
- A pasta de destino deve existir antes de iniciar o processamento.

---

## Contribuição

Contribuições são bem-vindas! Para colaborar:

1. Faça um **fork** do repositório
2. Crie uma branch para sua feature (`git checkout -b feature/nova-funcionalidade`)
3. Commit suas alterações (`git commit -m 'Adiciona nova funcionalidade'`)
4. Push para a branch (`git push origin feature/nova-funcionalidade`)
5. Abra um **Pull Request**

Para reportar bugs ou sugerir melhorias, abra uma [issue](https://github.com/seu-usuario/E-Fal/issues).

---

## Licença

Distribuído sob a licença **MIT**. Consulte o arquivo [LICENSE](LICENSE) para mais detalhes.

---

<div align="center">

Desenvolvido para automação jurídica no **TJRJ**

</div>

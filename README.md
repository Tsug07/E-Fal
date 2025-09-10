# E-Fal - Emissão de Certidão de Falência

## Descrição
O **E-Fal** é uma ferramenta de automação desenvolvida para facilitar a emissão de certidões de falência no site do Tribunal de Justiça do Rio de Janeiro (TJRJ). A aplicação utiliza uma interface gráfica (GUI) baseada em Tkinter para interagir com o usuário, permitindo a validação de arquivos Excel, o processamento de pedidos de certidões e o download automático dos arquivos PDF gerados, salvos em uma pasta especificada. A automação é realizada com o Selenium WebDriver, eliminando a necessidade de interação manual com a interface do navegador.

## Funcionalidades
- **Validação de Excel**: Verifica se o arquivo Excel fornecido possui as colunas esperadas ("Código", "Cliente", "CND", "Pedido") e carrega os dados em uma interface de visualização.
- **Processamento Automático**: Preenche formulários no site do TJRJ, acessa as certidões e realiza o download dos arquivos PDF.
- **Download Gerenciado**: Configura o Chrome para salvar os arquivos diretamente na pasta de destino, renomeando-os automaticamente com base no código do pedido e na data de vencimento.
- **Interface Gráfica**: Interface intuitiva com campos para selecionar o arquivo Excel, pasta de destino, URL do TJRJ, barra de progresso, visualização de dados e log de execução.
- **Log de Execução**: Registra todas as ações em um arquivo de log com timestamp para acompanhamento e depuração.

## Requisitos
- **Sistema Operacional**: Windows (devido à configuração de caminhos e uso do Chrome).
- **Python**: Versão 3.6 ou superior.
- **Bibliotecas Python**:
  - `tkinter` (geralmente incluído com Python)
  - `openpyxl`
  - `selenium`
  - `webdriver_manager`
  - `psutil`
- **Navegador**: Google Chrome instalado.
- **Dependências Adicionais**: O `webdriver_manager` instala automaticamente o ChromeDriver compatível.

## Instalação
1. Clone ou baixe o repositório contendo o arquivo `testeCla.py`.
2. Instale as dependências necessárias:
   ```bash
   pip install openpyxl selenium webdriver_manager psutil
   ```
3. Certifique-se de que o Google Chrome está instalado no sistema.
4. Crie uma pasta para armazenar os logs (ex.: `TJRJ_Logs`) e os arquivos baixados (ex.: `C:\Users\VM001\Documents\CNDs\FALENCIA`).

## Como Usar
1. Execute o script `testeCla.py`:
   ```bash
   python e-fal.py
   ```
2. Na interface gráfica:
   - **Selecione o arquivo Excel**: Clique em "Procurar" no campo "Arquivo Excel" e escolha um arquivo `.xlsx` com as colunas "Código", "Cliente", "CND" e "Pedido".
   - **Defina a pasta de destino**: Escolha a pasta onde os arquivos PDF serão salvos (padrão: `C:\Users\VM001\Documents\CNDs\FALENCIA`).
   - **Valide o Excel**: Clique em "Validar Excel" para verificar o formato do arquivo e visualizar os dados na tabela.
   - **Inicie o processamento**: Clique em "Iniciar Processamento" para começar a automação.
   - **Pare o processamento**: Caso necessário, clique em "Parar" para interromper a execução.
3. Acompanhe o progresso no log de execução e na barra de progresso.

## Formato do Excel
O arquivo Excel deve ter as seguintes colunas na primeira linha:
- **Código**: Identificador único do pedido.
- **Cliente**: Nome do cliente.
- **CND**: Campo opcional (pode estar vazio).
- **Pedido**: Número do pedido a ser consultado no site do TJRJ.

Exemplo:
| Código | Cliente         | CND             | Pedido       |
|--------|-----------------|-----------------|--------------|
| 12345  | João Silva      | CND FALENCIA   | 2023-0012345 |
| 67890  | Maria Oliveira  | CND FALENCIA   | 2023-0067890 |

## Configurações
- **URL do TJRJ**: Por padrão, a URL é `https://www3.tjrj.jus.br/CJE/certidao/judicial/visualizar?modelo=visualizar`. Pode ser alterada na interface, se necessário.
- **Pasta de Logs**: Os logs são salvos em `TJRJ_Logs` no mesmo diretório do script, com nomes no formato `tjrj_log_AAAAMMDD_HHMMSS.txt`.
- **Perfil do Chrome**: Um perfil específico (`C:\PerfisChrome\automacao\Profile 1`) é usado para manter as configurações do navegador. Certifique-se de fazer login no site do TJRJ, se necessário, na primeira execução.

## Observações
- A automação depende do site do TJRJ estar acessível e manter a estrutura atual das páginas. Alterações no site podem exigir ajustes nos XPaths usados no código.
- Certifique-se de que a pasta de destino existe antes de iniciar o processamento.
- O programa cria automaticamente a pasta `TJRJ_Logs` para armazenar os logs, se ela não existir.
- A automação usa um perfil do Chrome para persistir sessões. Se o perfil não existir, será criado, e o usuário pode precisar fazer login manualmente na primeira execução.

## Limitações
- A ferramenta foi projetada para o site específico do TJRJ e pode não funcionar em outros sistemas sem adaptações.
- O download automático depende das configurações do Chrome. Problemas com permissões ou configurações de segurança do navegador podem afetar o salvamento dos arquivos.
- Requer conexão estável com a internet durante o processamento.

## Contribuições
Contribuições são bem-vindas! Para sugerir melhorias ou reportar problemas, crie uma issue no repositório ou envie um pull request com as alterações propostas.

## Licença
Este projeto é distribuído sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes (se aplicável).
# Sistema de Gerenciamento de Manutenções da Academia

Este projeto Python oferece uma ferramenta simples e eficiente para registrar e acompanhar manutenções preventivas e corretivas dos aparelhos de uma academia, armazenando os dados em uma planilha Excel organizada.

## 🚀 Funcionalidades

* Registro de Manutenção Preventiva Geral: Preenche automaticamente uma linha para cada aparelho na planilha mestra e em suas abas individuais.
* Cálculo automático da próxima manutenção preventiva (quinzenalmente).
* Registro de Manutenção Corretiva: Permite selecionar um aparelho específico, registrar a correção e detalhar a observação obrigatória.
* Cálculo da "Próxima Manutenção Prevista" para corretivas: Exibe a data da próxima preventiva agendada para aquele aparelho, baseada na última preventiva registrada.
* Organização em abas: Cria uma aba exclusiva para cada aparelho (`NomeAparelho (Patrimônio)`) com seu histórico de manutenções.
* Planilha Mestra: Mantém um resumo consolidado de todas as manutenções realizadas.
* Validação de entrada: Garante que apenas aparelhos pré-cadastrados sejam selecionados.

## 🛠️ Tecnologias Utilizadas

* **Linguagem:** Python 3.9+
* **Bibliotecas:**
    * `openpyxl` (para manipulação de arquivos Excel)
    * `datetime` (módulo padrão do Python para lidar com datas e horas)
    * `os` (módulo padrão do Python para interações com o sistema operacional)

## ⚙️ Como Usar

### Pré-requisitos

* Python 3.8 ou superior
* `pip` (gerenciador de pacotes do Python)

### Instalação

1.  Baixe o arquivo `registro_manutencao.py` para o seu computador.
2.  Abra o terminal (Prompt de Comando no Windows, Terminal no macOS/Linux).
3.  Navegue até o diretório onde você salvou o arquivo:
    ```bash
    cd caminho/para/o/seu/diretorio
    ```
4.  Crie e ative um ambiente virtual (recomendado):
    ```bash
    python -m venv venv
    # No Windows: .\venv\Scripts\activate
    # No macOS/Linux: source venv/bin/activate
    ```
5.  Instale a biblioteca `openpyxl`:
    ```bash
    pip install openpyxl
    ```

### Execução

1.  Com o ambiente virtual ativado, execute o script:
    ```bash
    python registro_manutencao.py
    ```
2.  Siga as instruções no terminal para escolher o tipo de manutenção (Preventiva ou Corretiva) e fornecer os detalhes solicitados.
3.  Um arquivo `preventivas_academia.xlsx` será criado (ou atualizado) no mesmo diretório do script.

## 🧑‍💻 Autor(es)

* **Pyolho** (Seu parceiro de desenvolvimento)
* **[Rafael S. Pereira/rafacat]**

## 📄 Licença

Este projeto é de uso proprietário para fins de gerenciamento interno da academia.

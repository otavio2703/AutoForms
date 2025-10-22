# Automatizador Desktop com GUI (v3.6)

Este é um script Python que fornece uma interface gráfica (GUI) para automatizar tarefas repetitivas de desktop. Ele é projetado para ler dados de uma planilha Excel (como números de contrato) e, em seguida, usar reconhecimento de imagem (OpenCV) e controle de mouse/teclado (`pydirectinput`) para processar cada item em um aplicativo de desktop.

A principal característica desta automação é o uso de reconhecimento de imagem em vez de coordenadas fixas (X, Y), tornando-a mais robusta a pequenas mudanças de layout, resoluções de tela diferentes ou janelas em posições diferentes.


## Funcionalidades Principais

* **Interface Gráfica Amigável:** Criada com `tkinter` para facilitar a configuração dos arquivos e imagens.
* **Leitura de Excel:** Carrega dados de arquivos `.xlsx` ou `.xls` usando `pandas`.
* **Contagem de Itens:** Verifica e informa quantos contratos (linhas) serão processados.
* **Automação Robusta:** Usa `OpenCV` para encontrar elementos na tela em tempo real.
* **Lógica de Botão Duplo (v3.6):** Capaz de reconhecer um botão em dois estados diferentes (ex: estado **nativo** e estado **hover**), tornando a automação mais confiável contra elementos dinâmicos.
* **Controle Seguro:** Utiliza `pydirectinput`, que é mais eficaz em aplicativos (especialmente jogos ou ambientes virtualizados) que podem bloquear controles de automação padrão.
* **Log de Status:** Exibe um log em tempo real do que está acontecendo.
* **Controle de Parada:** Permite que o usuário interrompa a automação com segurança.

## Tecnologias Utilizadas

* [Python 3](https://www.python.org/)
* [Tkinter](https://docs.python.org/3/library/tkinter.html) (para a Interface Gráfica)
* [Pandas](https://pandas.pydata.org/) (para leitura de arquivos Excel)
* [OpenCV-Python](https://pypi.org/project/opencv-python/) (para reconhecimento de imagem)
* [Pillow (PIL)](https://pypi.org/project/Pillow/) (para captura de tela)
* [PyDirectInput](https://pypi.org/project/PyDirectInput/) (para controle de mouse e teclado)

## Instalação

1.  **Clone o repositório** (ou baixe o arquivo `.py`):
    ```bash
    git clone https://github.com/otavio2703/AutoForms
    cd autoforms
    ```

2.  **Crie um ambiente virtual** (altamente recomendado):
    ```bash
    # Windows
    python -m venv venv
    venv\Scripts\activate

    # macOS/Linux
    python3 -m venv venv
    source venv/bin/activate
    ```

3.  **Instale as dependências necessárias:**
    ```bash
    pip install pandas opencv-python pillow pydirectinput
    ```

## Como Usar

Para que a automação funcione, você precisa "ensinar" ao script o que procurar. Isso é feito fornecendo imagens de referência (templates).

### 1. Preparação: As Imagens

Antes de executar o script, você deve tirar capturas de tela de cada elemento interativo.

**Dicas Cruciais para Capturas de Tela:**

* **Sejam Únicas:** A imagem deve ser a menor porção única do elemento. Por exemplo, em vez de tirar print do botão inteiro com bordas, capture apenas o texto ou o ícone dentro dele.
* **Não Capture o Mouse:** O cursor do mouse não pode aparecer na captura de tela.
* **Consistência:** A aparência do elemento na tela deve ser *exatamente* igual à da imagem de referência.
* **Formato:** Salve todas as imagens como `.png`.

Você precisará das 5 imagens a seguir:
1.  **Campo de Contrato:** Uma imagem do rótulo ou ícone do campo onde o número do contrato será digitado.
2.  **Botão Pesquisar:** O botão que inicia a busca.
3.  **Botão Crédito (Nativo):** O botão "Crédito Recebido" em seu estado normal, sem o mouse em cima.
4.  **Botão Crédito (Hover):** O botão "Crédito Recebido" quando o mouse está sobre ele (se ele mudar de cor, brilho ou estilo).
5.  **Botão Final:** O botão de confirmação que aparece após rolar a página.

### 2. Executando a Automação

1.  Execute o script:
    ```bash
    python script v3.5 copy.py
    ```
2.  **Passo 1: Configuração do Excel**
    * Clique em "Procurar..." e selecione seu arquivo Excel.
    * No campo "Nome da Coluna", digite o nome *exato* da coluna que contém os números dos contratos (o padrão é "Numero do Contrato").
    * A interface atualizará o "Total de Contratos a processar".

3.  **Passo 2: Configuração das Imagens**
    * Clique em "Procurar..." para cada um dos 5 campos de imagem e selecione os arquivos `.png` que você preparou na etapa anterior.

4.  **Passo 3: Iniciar**
    * O botão "Iniciar Automação" ficará verde quando todas as configurações estiverem prontas.
    * Clique no botão. Você terá **5 segundos** para mudar o foco para a janela do aplicativo que você deseja automatizar.

5.  **Acompanhamento**
    * O script começará a processar o primeiro contrato.
    * A caixa de "Status e Log" mostrará cada etapa: "Procurando...", "Encontrado...", "Digitando...".
    * Em caso de erro (ex: não encontrar uma imagem após 10 segundos), ele registrará a falha e passará para o próximo contrato.

6.  **Parando a Automação**
    * A qualquer momento, você pode clicar no botão "Parar Automação".
    * O script não vai parar *imediatamente*. Ele terminará o contrato que está processando no momento e, em seguida, parará de forma limpa, sem iniciar o próximo.

## Detalhe da v3.6: Lógica de Busca Dupla

O problema mais comum em automação de cliques é que, ao mover o mouse para um botão, a aparência do botão muda (estado *hover*). Se o script estiver procurando apenas pela imagem nativa, ele falhará no momento exato em que o mouse chegar ao destino.

A função `encontrar_e_clicar_duas_imagens` resolve isso:

1.  O script tira uma foto da tela.
2.  Ele procura pela imagem **Nativa** (`imagem_botao_liquidacao`).
3.  Se encontrar, clica e retorna sucesso.
4.  Se *não* encontrar, ele procura pela imagem **Hover** (`imagem_botao_liquidacao_hover`) na *mesma foto da tela*.
5.  Se encontrar a imagem hover, clica e retorna sucesso.
6.  Se não encontrar nenhuma das duas, ele espera meio segundo e tenta o processo todo novamente, até o limite de tempo (10s).

Isso garante que, não importa se o mouse já está sobre o botão ou não, o script será capaz de localizá-lo e clicar.

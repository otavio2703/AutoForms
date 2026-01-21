
# Automatizador Desktop com GUI — v3.13

> **Atualização**: Este README documenta a versão **v3.13** do script, com geração de relatórios em **Excel** (coluna *Status*) e **TXT** com a lista de contratos que falharam. _(Atualizado em 21/01/2026)_

## Sumário
- [Visão Geral](#visão-geral)
- [Novidades na v3.13](#novidades-na-v313)
- [Funcionalidades Principais](#funcionalidades-principais)
- [Requisitos](#requisitos)
- [Instalação](#instalação)
- [Como Usar](#como-usar)
- [Preparação das Imagens (Templates)](#preparação-das-imagens-templates)
- [Fluxo de Processamento](#fluxo-de-processamento)
- [Saídas e Relatórios](#saídas-e-relatórios)
- [Configurações e Ajustes](#configurações-e-ajustes)
- [Boas Práticas para Reconhecimento por Imagem](#boas-práticas-para-reconhecimento-por-imagem)
- [Solução de Problemas (FAQ)](#solução-de-problemas-faq)
- [Changelog](#changelog)
- [Avisos](#avisos)

---

## Visão Geral
Script Python com **interface gráfica (Tkinter)** para automatizar tarefas em aplicações desktop, lendo uma planilha Excel (ex.: números de contrato) e utilizando **reconhecimento de imagem (OpenCV)** junto a **controle de mouse/teclado** via `pydirectinput`. A abordagem por *template matching* torna a automação **menos dependente de coordenadas X/Y fixas**, funcionando melhor diante de pequenas mudanças de layout e posição de janelas.

## Novidades na v3.13
- **Nova coluna `Status` no Excel** marcando **Sucesso** (Liquidado) ou **Erro** por contrato.
- **Geração de arquivo Excel processado** com sufixo `*_PROCESSADO.xlsx` (para não sobrescrever o original).
- **Geração de arquivo `.txt`** contendo **somente os contratos com erro**, com *timestamp* no nome.

## Funcionalidades Principais
- **GUI amigável (Tkinter):** seleção do Excel, definição da coluna de contratos, escolha dos templates (`.png`) e botões de controle (Iniciar/Parar) com barra de progresso e log em tempo real.
- **Leitura de Excel (Pandas):** suporta `.xlsx`/`.xls`; a coluna padrão sugerida é **"Numero do Contrato"** (ajustável na interface).
- **Contagem de itens:** exibe o total de linhas válidas (não nulas) a processar.
- **Automação robusta:** `OpenCV` + *template matching*; clique por **âncora + offset** para campos de digitação; procura por **múltiplos estados de botão** (Nativo/Hover/Pequeno) para maior confiabilidade.
- **Controles seguros:** `pydirectinput` para mover/clicar e enviar teclas (inclusive sequências de `Enter` e rolagem por setas).
- **Logs e progresso:** área de log com tudo que acontece e rótulo de progresso (ex.: “Processando 3 de 20…”).
- **Parada segura:** o botão **Parar Automação** encerra o ciclo com segurança.

## Requisitos
- **SO:** Windows (recomendado). Outros SOs podem funcionar, mas não são o alvo principal.
- **Python:** 3.8+ (recomendado 3.10+)
- **Bibliotecas:**
  - `pandas` (e `openpyxl` para `.xlsx`)
  - `opencv-python`
  - `pillow`
  - `pydirectinput`

## Instalação

1) **Clone o repositório (ou baixe o .py):**
```bash
git clone https://github.com/otavio2703/AutoForms
cd AutoForms
```

2) **Crie e ative um ambiente virtual (opcional, recomendado):**
```bash
# Windows
python -m venv venv
venv\Scriptsctivate

# macOS/Linux
python3 -m venv venv
source venv/bin/activate
```

3) **Instale as dependências:**
```bash
pip install pandas openpyxl opencv-python pillow pydirectinput
```

## Como Usar
1. **Abra o script** (`python script13.3.py`).
2. **Selecione o Excel** (botão *Procurar*). Informe/ajuste a **coluna de contratos** (padrão sugerido: `Numero do Contrato`). O app exibirá o **total de contratos** válidos.
3. **Aponte os templates (.png)** na seção “Configuração das Imagens”: 
   - **Rótulo do Campo (âncora)** + defina o **offset (X, Y)** do clique para alcançar o campo de digitação.
   - **Botão Pesquisar**.
   - **Crédito Recebido** em **3 variações**: *Nativo*, *Hover* e *Pequeno*.
   - **Botão Final (Confirmar)**.
4. **Clique em Iniciar Automação**. Observe o **log** e o **progresso**. Use **Parar Automação** para interromper com segurança.

> **Importante:** durante a execução, **não mova o mouse** nem interaja com a janela alvo para evitar interferências nos cliques.

## Preparação das Imagens (Templates)
**Dicas cruciais:**
- **Seja específico:** capture a **menor área única** do elemento (evite bordas e áreas vazias). 
- **Sem cursor:** garanta que **o mouse não apareça** no recorte.
- **Mesma aparência:** a imagem de referência deve estar **idêntica** ao estado exibido no app (cor, fonte, zoom, idioma, tema, DPI etc.).
- **Formato:** salve como **`.png`**.

**Checklist de imagens**:
1. Rótulo/ícone do **campo de contrato** (âncora de clique + offset).
2. **Botão Pesquisar**.
3. **Crédito Recebido (Nativo)**.
4. **Crédito Recebido (Hover)**.
5. **Crédito Recebido (Pequeno)**.
6. **Botão Final (Confirmar)**.

## Fluxo de Processamento
1. Encontrar **rótulo/âncora** e **clicar com offset** para focar o campo de contrato.
2. **Colar** o número do contrato (clipboard + `Ctrl+V`).
3. Clicar no **Botão Pesquisar**.
4. Localizar e clicar em **Crédito Recebido** (tenta **Nativo → Hover → Pequeno** até achar).
5. **Rolar a tela para baixo** (sequência de `ArrowDown`).
6. Clicar no **Botão Final (Confirmar)**.
7. Confirmar com **`Enter` → aguarda 2s → `Enter`**.
8. Voltar ao **campo de contrato**, **selecionar tudo** (`Ctrl+A`) e **limpar** (`Backspace`).
9. Registrar **Status** (Sucesso/Erro) e seguir para o próximo.

## Saídas e Relatórios
- **Excel processado:** o arquivo de origem é salvo como um **novo** arquivo com sufixo `*_PROCESSADO.xlsx`, adicionando/atualizando a coluna **`Status`** por linha.
- **TXT de erros:** gerado ao final **somente** se houver falhas, nomeado como `erros_log_YYYYMMDD_HHMMSS.txt` na **mesma pasta do Excel**.

> Valores de **Status**: `Liquidado` (sucesso) e `Erro` (falha). Revise a ortografia caso personalize.

## Configurações e Ajustes
- **Coluna de contratos:** campo editável na GUI (padrão: `Numero do Contrato`).
- **Offset (X, Y) do clique:** define o deslocamento a partir do **rótulo/âncora** para alcançar o campo. Padrão sugerido no app: **X=200, Y=0** (ajuste conforme seu layout).
- **Confiança do template matching:** fixada em **0.8**.
- **Timeout de busca por elemento:** **10s** (com tentativas a cada ~0,5s).

## Boas Práticas para Reconhecimento por Imagem
- **DPI/escala** do sistema: prefira **100%** (sem zoom) para consistência.
- **Iluminação/tema:** mantenha o **mesmo tema** (claro/escuro) entre as capturas e a execução.
- **Idioma e fonte:** diferença de idioma/fonte/antialiasing pode reduzir a confiança.
- **Janelas/posição:** garanta que a **mesma tela** esteja visível e desbloqueada.

## Solução de Problemas (FAQ)
**Não encontra o elemento mesmo com a imagem correta.**  
• Tente **recortar uma área menor e mais única**.  
• Ajuste o **offset** (para o campo) e **reposicione** a janela para exibir o elemento por completo.  
• Verifique **DPI/escala** e **tema**. 

**Clica no lugar errado do campo.**  
• Revise o **offset X/Y** até centralizar no **ponto correto** do campo. 

**Excel não salva.**  
• Feche o arquivo original se estiver aberto. O script salva um **novo** arquivo `*_PROCESSADO.xlsx`. 

**Planilha não carrega a coluna.**  
• Confira o **nome exato** da coluna (sensível a acentuação e espaços). 

## Changelog
- **v3.13**
  - Coluna **`Status`** no Excel (por linha de contrato).
  - Salva Excel com sufixo `*_PROCESSADO.xlsx`.
  - Gera **TXT** com lista de contratos que falharam.

## Avisos
- **Uso responsável:** respeite as políticas da aplicação alvo. 
- **Interferências** (mover o mouse/teclado) podem causar falhas durante a automação. 
- Este projeto é fornecido "como está", sem garantias.

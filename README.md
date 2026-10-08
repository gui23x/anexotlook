# Anexotlook

<img src="https://i.pinimg.com/originals/ad/4d/77/ad4d77597768b41307194c1e3542a284.jpg" alt="Anexotlook" />

Aplicação desktop para extrair anexos de e-mails do Microsoft Outlook de forma local e orientada por interface gráfica. O projeto foi migrado para o gerenciamento de dependências com `uv`, simplificando a configuração do ambiente e a execução do aplicativo em Windows.

## Visão geral

O Anexotlook conecta-se ao Outlook instalado no computador, lista as contas e pastas disponíveis, sincroniza os itens recentes e permite selecionar anexos específicos para exportação em massa. A interface foi construída com PySide6 e QFluentWidgets, com visual inspirada no Fluent Design da Microsoft.

### Funcionalidades principais

- Listagem de contas e pastas do Outlook
- Sincronização assíncrona de e-mails em thread separada
- Visualização de anexos com seleção individual ou por lote
- Download dos anexos em diretório escolhido pelo usuário
- Ajuste automático da cor de destaque do Windows na interface
- Compatibilidade com integração local via COM/MAPI do Outlook

---

## Requisitos do sistema

| Requisito                   | Especificação                                        |
| :-------------------------- | :--------------------------------------------------- |
| Sistema operacional         | Windows 10/11                                        |
| Outlook                     | Microsoft Outlook instalado e autenticado localmente |
| Python                      | 3.12+                                                |
| Gerenciador de dependências | `uv`                                                 |

> O uso de COM/MAPI e acesso ao registro do Windows exige execução no ambiente Windows. O projeto não foi projetado para funcionar em Linux/macOS como ambiente principal.

---

## Dependências do projeto

O projeto utiliza as bibliotecas abaixo, declaradas em `pyproject.toml`:

| Biblioteca               | Finalidade                                                    |
| :----------------------- | :------------------------------------------------------------ |
| `pyside6`                | Framework principal da interface gráfica                      |
| `pyside6-fluent-widgets` | Componentes visuais no estilo Fluent Design                   |
| `pywin32`                | Integração com Windows e Outlook via `win32com` e `pythoncom` |

---

## Configuração com uv

### 1) Instalar o uv

Se ainda não estiver instalado:

```bash
pip install uv
```

Ou siga a instalação oficial em: https://docs.astral.sh/uv/

### 2) Instalar dependências do projeto

Na raiz do repositório, execute:

```bash
uv sync
```

Esse comando cria o ambiente virtual do projeto e instala as dependências definidas no `pyproject.toml` com base no `uv.lock`.

### 3) Executar a aplicação

```bash
uv run python src/main.py
```

Ou, caso o ambiente esteja ativado:

```bash
python src/main.py
```

---

## Estrutura do projeto

```text
anexotlook/
├── README.md
├── pyproject.toml
├── uv.lock
├── src/
│   └── main.py
└── .python-version
```

A aplicação principal está em `src/main.py` e utiliza as dependências gerenciadas pelo `uv`.

---

## Fluxo de uso

1. Abra o aplicativo.
2. Selecione o perfil do Outlook no campo **CONTA**.
3. Escolha a pasta de e-mails desejada no campo **PASTA**.
4. Clique em **Sincronizar** para carregar os itens mais recentes.
5. Marque os anexos desejados na tabela.
6. Use **Selecionar Todos** para selecionar lote, quando necessário.
7. Clique em **Baixar Selecionados** e escolha a pasta de destino.

---

## Detalhes de arquitetura

- **Gerenciamento de concorrência**: a sincronização de caixas de correio acontece em uma thread separada (`SyncWorker`). A thread usa `pythoncom.CoInitialize()` para evitar travamentos na interface principal enquanto cria um COM apartment isolado para leitura do Outlook.
- **Recuperação de dados**: o sistema acessa o namespace MAPI do Outlook, lista contas e subpastas disponíveis, ordena itens por `ReceivedTime` e limita a leitura aos 250 itens mais recentes para manter o desempenho.
- **Integração com Windows**: a aplicação lê a cor de destaque do sistema em `HKEY_CURRENT_USER\Software\Microsoft\Windows\DWM` e converte o valor para RGB usado no Qt.
- **Sanitização de arquivos**: nomes de anexos são tratados com expressão regular para remover caracteres inválidos e evitar falhas durante a gravação no disco.

---

## Observações importantes

- Certifique-se de que o Outlook esteja configurado com pelo menos um perfil ativo.
- Verifique se o arquivo `mail.ico` está presente no diretório de trabalho antes da execução, pois ele é usado para os ícones da aplicação.
- Para desenvolvimento e execução local, prefira usar `uv` em vez de `pip install` manual, para manter o ambiente consistente com as versões declaradas no projeto.

---

## Desenvolvimento

Para atualizar dependências do projeto:

```bash
uv add <pacote>
```

Para atualizar o lock file após mudanças:

```bash
uv lock
```

Para rodar verificações locais ou scripts extras, use sempre `uv run` para garantir que o ambiente correto do projeto seja utilizado.

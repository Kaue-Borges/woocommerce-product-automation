# WooCommerce Product Automation

Ferramentas em **Python, Flask e Playwright** para auxiliar a atualização de produtos existentes no WooCommerce, principalmente na organização e no envio de **imagem principal, galeria de fotos e arquivos para download**.

> **Situação:** scripts de automação em evolução. O repositório reúne diferentes versões da interface de cadastro e uma rotina separada para remoção de mídias. Não representa ainda uma plataforma única com API ou importação CSV implementada.

## Funcionalidades presentes no código

- Localizar produtos no painel WooCommerce pelo **SKU**.
- Encontrar pastas locais de produtos e classificar arquivos para envio.
- Cadastrar/atualizar **imagem destacada**, **galeria** e **arquivos para download**.
- Processar vários SKUs e registrar progresso, resultados e erros.
- Usar painéis locais em Flask para selecionar uma planilha, iniciar lotes e consultar status ou relatório.
- Trabalhar com configurações de mais de uma loja na versão `cadastro_dashboardd.py`.
- Remover mídias e downloads de produtos selecionados, usando os scripts independentes em `Remover_fotos/`.

**Atenção:** a rotina de remoção modifica dados no WooCommerce. Use somente com autorização, testes e backup adequado. A automação de cadastro também altera produtos reais.

## Estrutura atual

```text
.
├── cadastro_choice.py
├── cadastro_dashboard.py
├── cadastro_dashboardd.py
├── Remover_fotos/
│   ├── remover_dashboard.py
│   └── remover_tudo_woocommerce.py
├── env.example
├── .gitignore
└── README.md
```

| Arquivo | Responsabilidade |
| --- | --- |
| `cadastro_choice.py` | Automação interativa/por lote via terminal e Playwright |
| `cadastro_dashboard.py` | Dashboard Flask para cadastro de produtos |
| `cadastro_dashboardd.py` | Versão ampliada do dashboard com seleção de loja |
| `Remover_fotos/remover_dashboard.py` | Interface local para operações de remoção |
| `Remover_fotos/remover_tudo_woocommerce.py` | Rotinas de busca por SKU e remoção de mídias |

Os dashboards possuem implementação parcialmente duplicada; uma futura refatoração poderá consolidar os fluxos sem perder as funções existentes.

## Requisitos

- **Python 3.10+** (o código usa anotações de tipo com `|`).
- **Google Chrome** disponível para as automações de navegador.
- Acesso autorizado ao painel WordPress/WooCommerce.
- Pastas de mídia organizadas por produto e uma lista de SKUs para os fluxos em lote.

Não há um `requirements.txt` consolidado nesta versão. Para testar em um ambiente virtual:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install flask playwright
```

Esses pacotes cobrem as importações externas centrais dos scripts inspecionados. Dependendo do ambiente, a configuração do navegador e outras dependências podem exigir ajustes.

## Execução

**Cadastro pelo terminal:**

```powershell
python cadastro_choice.py --help
```

O script aceita parâmetros como `--base-url`, `--root` e `--headless`. Confira os valores antes de executar.

**Dashboard de cadastro:**

```powershell
python cadastro_dashboardd.py
```

A interface principal dessa versão inicia em `http://127.0.0.1:5056`. A versão `cadastro_dashboard.py` também utiliza essa porta: execute **apenas uma das duas por vez**.

**Dashboard de remoção de mídias:**

```powershell
python Remover_fotos/remover_dashboard.py
```

A interface utiliza `http://127.0.0.1:8788`. **Não faça testes de remoção em uma loja de produção.**

## Configurações e segurança

Os scripts ainda contêm URLs de lojas e caminhos locais fixos que precisam ser revisados ao mudar de computador ou ambiente. O arquivo [`env.example`](env.example) é apenas **ilustrativo**: as variáveis listadas nele não substituem automaticamente as configurações usadas pelos scripts.

O dashboard ampliado também prevê autenticação opcional por um arquivo local `credenciais.json`. **Nunca versione credenciais, cookies, perfis Chrome, dados de produtos internos ou capturas com informações privadas.**

O repositório é público. Antes de publicar novas alterações, revise os dados e o código com atenção e obtenha as permissões necessárias.

## Próximas melhorias

- [ ] Centralizar as configurações de lojas, diretórios e autenticação fora do código.
- [ ] Consolidar as duas versões do dashboard.
- [ ] Criar `requirements.txt` e testes com mocks, sem acessar lojas reais.
- [ ] Documentar exemplos fictícios de planilha e organização de pastas.
- [ ] Separar com mais clareza ações de cadastro e ações destrutivas.
- [ ] Adicionar validações e proteção contra alterações indesejadas.

## Autoria

**Kaue Borges** — projeto de automação de processos e integração com WooCommerce.

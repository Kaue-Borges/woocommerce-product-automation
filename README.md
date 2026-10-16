# Automação de Cadastro de Produtos no WooCommerce

Automação em Python para apoiar a digitalização e a preparação de cadastros de produtos para lojas WooCommerce.

> **Status:** em desenvolvimento. Este README descreve o objetivo e o fluxo planejado; as funcionalidades disponíveis devem ser confirmadas no código antes da publicação.

## Objetivo

Reduzir o trabalho manual no cadastro de produtos, padronizando informações e organizando materiais como fotos e fichas técnicas.

## Fluxo de trabalho

1. Selecionar a empresa/loja de destino (Choice, SAG ou Portwell).
2. Informar o código do produto (CIGAM), nome e dados complementares.
3. Anexar imagens e, quando disponível, o datasheet.
4. Preparar descrições e demais campos do cadastro.
5. Organizar os arquivos do produto em uma pasta identificada pelo código.
6. Gerar um CSV compatível com a importação do WooCommerce.
7. Importar os produtos e validar imagens, documentos e categorias na loja de destino.

**Importante:** algumas etapas podem estar em implementação ou depender de configuração manual. Atualize a lista após validar o comportamento real da aplicação.

## Tecnologias

- Python
- FastAPI
- Jinja2
- HTML, CSS e JavaScript
- WooCommerce (importação CSV / integração conforme a implementação)

## Estrutura do projeto

A estrutura das pastas será documentada após a revisão do código-fonte, evitando indicar arquivos ou módulos que ainda não existam.

## Como executar localmente

> Ajustar os comandos ao ponto de entrada e às dependências reais do projeto.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Se a aplicação utilizar Uvicorn com um objeto `app` em `main.py`, o comando usual é:

```powershell
uvicorn main:app --reload
```

## Configuração

Nunca envie credenciais de lojas, tokens de IA ou dados privados para o repositório. Use variáveis de ambiente e mantenha somente exemplos sem valores reais.

Consulte `.env.example` para organizar as configurações necessárias; ajuste os nomes às variáveis que seu código efetivamente utiliza.

## Segurança e privacidade

- Não versionar `.env`, senhas, tokens ou chaves de API.
- Não publicar bases de produtos, planilhas internas, fichas técnicas proprietárias ou arquivos de clientes sem autorização.
- Revisar a lista de arquivos antes de cada commit.
- Preferir repositório privado enquanto o projeto utilizar dados ou lógica interna das empresas.

## Próximos passos

- [ ] Revisar a estrutura atual e atualizar a documentação de instalação.
- [ ] Confirmar quais fluxos já estão funcionando nas três empresas.
- [ ] Adicionar exemplos de CSV com dados fictícios.
- [ ] Documentar erros comuns, testes e processo de importação.
- [ ] Incluir imagens de demonstração sem dados internos.

## Autoria

Desenvolvimento e documentação do projeto de automação de cadastros WooCommerce.

## Licença

Ainda não definida. Não adicionar licença aberta sem confirmar os direitos sobre o código e a autorização de divulgação.

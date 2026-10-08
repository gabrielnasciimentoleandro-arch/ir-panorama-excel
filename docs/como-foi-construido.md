# Como o IR Panorama foi construído

## 1. Leitura do desafio

O desafio solicita uma solução em Excel para reunir informações essenciais da declaração de Imposto de Renda, com menus de navegação, validações automáticas, links rápidos e documentação no GitHub.

## 2. Decisão autoral

Em vez de reproduzir as telas do material de aula, o projeto foi estruturado como um painel documental chamado **IR Panorama**. A proposta separa três camadas:

1. **cadastros:** titular e dependentes;
2. **registros financeiros:** rendimentos, bens e pagamentos;
3. **controle:** informes, notas, prazos e progresso.

A identidade visual, os textos, os exemplos, os valores e o fluxo foram criados para este projeto.

## 3. Estrutura da planilha

A pasta de trabalho possui dez abas:

- Início;
- Titular;
- Dependentes;
- Rendimentos;
- Bens;
- Pagamentos;
- Informes;
- Notas;
- Guia;
- Apoio.

A aba Apoio concentra listas e permanece oculta por padrão para não atrapalhar o uso diário.

## 4. Recursos empregados

- tabelas estruturadas com filtros;
- validação por listas;
- intervalos nomeados;
- SOMA, CONT.SE, CONT.SES, SE, SEERRO, CONT.VALORES, MÁXIMO e HOJE;
- formatação condicional;
- links internos entre abas;
- links externos somente para portais oficiais;
- gráfico de rosca para status documental;
- cache OOXML para visualizadores que não calculam fórmulas;
- metadados de autoria no arquivo.

## 5. Cenário demonstrativo

O arquivo público usa nomes como “Marina Exemplo”, documentos compostos por zeros e instituições fictícias. Os valores existem apenas para demonstrar os cálculos e não formam uma declaração fiscal válida.

## 6. Segurança

A solução não possui macro, script embutido, API, conexão bancária nem sincronização automática. O preenchimento ocorre localmente. A documentação orienta a manter a cópia real fora do repositório público.

## 7. Testes

O script `scripts/verificar_planilha.py` testa estrutura, fórmulas, caches, tabelas, listas, imagens, links, autoria e ausência de vínculos externos. O resultado atual está em `docs/validacao.txt`.

## 8. Autoria

**0Barone (@gabrielnasciimentoleandro-arch)**

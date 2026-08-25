# Painel Yubico

Sistema de gestão operacional de dispositivos **Yubico**, desenvolvido em **Google Apps Script** com interface web em HTML. O projeto centraliza indicadores de colaboradores, controle de inventário, organização de formulários no Google Drive, exportações filtradas, empréstimos, devoluções, informativos manuais e uma tela de TV para acompanhamento visual.

> O sistema utiliza uma planilha Google Sheets como base operacional e uma pasta do Google Drive como repositório de documentos. As novas abas de controle são criadas automaticamente quando necessário.

## Visão geral

O painel foi estruturado para apoiar a operação diária de Yubicos sem exigir banco de dados externo. A aba `Resumo_HC` é a fonte principal das informações de colaboradores e dispositivos, enquanto o Google Drive armazena formulários e documentos enviados. O `Code.gs` concentra a regra de negócio e o acesso aos serviços Google; o `Index.html` concentra a interface principal.

| Componente | Responsabilidade |
| --- | --- |
| Google Sheets | Fonte de dados de colaboradores, inventário, relatórios e históricos operacionais. |
| Google Drive | Estrutura de pastas e armazenamento de PDFs e formulários de empréstimo. |
| `Code.gs` | Backend em Apps Script, regras de negócio, cálculo de indicadores e operações de planilha/Drive. |
| `Index.html` | Painel web com Dashboard, Upload, Inventário, Exportações, Empréstimos e Informativos. |
| `TV.html` | Tela de acompanhamento para TV, incluindo o slide de Informativos. |

## Funcionalidades

| Módulo | Principais recursos |
| --- | --- |
| **Dashboard** | Mostra total de pessoas no HC, pessoas com Yubico, Yubicos sem formulário, pessoas sem Yubico com formulário e visão agrupada por empregador. |
| **Atualização do Drive** | Varre a pasta raiz configurada, atualiza o relatório de arquivos no Drive, mantém log de execução e armazena o último processamento. |
| **Upload de PDFs** | Permite selecionar uma pasta da árvore do Drive e salvar documentos PDF no destino escolhido. |
| **Inventário** | Registra as quantidades dos modelos Yubico 300, 325 e 600, mantém o histórico e mostra o saldo do último inventário. |
| **Exportações** | Filtra registros da `Resumo_HC` por empregador, departamento, turno, Yubico, formulário e conferência do empregador; exporta CSV e XLSX. |
| **Empréstimos** | Localiza o colaborador por CPF ou nome, registra empréstimos dos modelos 300, 325 ou 600, armazena o formulário no Drive e controla devoluções. |
| **Histórico de empréstimos** | Filtra por pendentes ou devolvidos, pesquisa registros, calcula dias pendentes e exporta os registros visíveis em CSV. |
| **Informativos** | Registra indicadores gerais e indicadores por empregador, incluindo pendências de devolução, saldo para desconto e desconto parcial. |
| **Tela TV** | Exibe dados operacionais em modo apresentação, incluindo Informativos no segundo slide. |

## Estrutura do repositório

```
.
├── Code.gs
├── Index.html
├── TV.html
├── README.md
├── testar_servidor_exportacao.js
├── testar_backend_emprestimos.js
├── testar_backend_informativos.js
├── testar_fluxo_filtros.js
├── testar_exportacao_emprestimos.js
└── validar_sintaxe.py
```

Os arquivos principais para a implantação são `Code.gs` e `Index.html`. O arquivo `TV.html` é necessário quando a funcionalidade de tela TV estiver habilitada. Os arquivos iniciados por `testar_` e `validar_` são utilizados apenas para validações locais e **não devem ser adicionados ao editor do Apps Script**.

## Pré-requisitos

| Requisito | Finalidade |
| --- | --- |
| Conta Google com acesso ao Google Sheets e Google Drive | Execução do Apps Script e acesso aos recursos usados pelo sistema. |
| Planilha Google Sheets operacional | Armazena a base `Resumo_HC` e as abas auxiliares. |
| Pasta raiz no Google Drive | Armazena documentos enviados e formulários de empréstimo. |
| Permissão de edição na planilha e na pasta | Necessária para criar abas, inserir registros e salvar arquivos. |
| Navegador atualizado | Acesso ao painel como Web App. |

O Apps Script solicita autorização antes de acessar dados privados da planilha e do Drive. Essa autorização deve ser concedida por uma conta que possua as permissões necessárias nos recursos configurados.[1]

## Configuração inicial

### 1. Criar ou abrir o projeto Apps Script

Na planilha Google Sheets, acesse **Extensões → Apps Script**. No editor, crie ou substitua o arquivo de script por `Code.gs` e adicione o arquivo HTML principal com o nome `Index`. Para usar a apresentação, adicione também um arquivo HTML chamado `TV` e copie o conteúdo de `TV.html`.

### 2. Atualizar a configuração do backend

No início do `Code.gs`, localize o objeto `CONFIG` e informe os identificadores da planilha e da pasta raiz do Drive.

```javascript
const CONFIG = {
  PLANILHA_ID: 'ID_DA_PLANILHA',
  ABA_RESUMO: 'Resumo_HC',
  ABA_RELATORIO_DRIVE: 'Relatorio_Drive',
  ABA_INVENTARIO: 'Inventario',
  ABA_BASE_HC: 'Base_HC',
  PASTA_ID: 'ID_DA_PASTA_RAIZ_DO_DRIVE'
};
```

Os IDs são encontrados na URL de cada recurso Google. Mantenha os nomes das abas configuradas ou altere tanto a configuração quanto os nomes reais na planilha de forma consistente.

> Não publique IDs reais, URLs internas de documentos ou dados pessoais em repositórios públicos.

### 3. Preparar a planilha de origem

A aba `Resumo_HC` deve possuir cabeçalho na primeira linha e dados a partir da segunda linha. As colunas abaixo são usadas diretamente pelos módulos de Exportações e Empréstimos.

| Coluna | Campo esperado | Uso no sistema |
| --- | --- | --- |
| A | Nome | Empréstimos e identificação do colaborador. |
| B | Departamento | Filtro de exportação e histórico de empréstimos. |
| C | CPF | Busca e validação do colaborador para empréstimos. |
| E | Função | Exibição no cadastro de empréstimo. |
| F | Nível | Exibição no cadastro de empréstimo. |
| G | Área | Exibição no cadastro de empréstimo. |
| H | Turno | Filtro de exportação e cadastro de empréstimo. |
| I | Setor | Cadastro e busca no histórico de empréstimos. |
| J | Empregador | Filtro de exportação e Informativos por empregador. |
| K | Yubico | Campo preenchido é tratado como **Sim** nas exportações. |
| L | Formulário | Valores **Sim** ou **Não** nas exportações. |
| O | Conferência | Valores **Sim** ou **Não** para validação do empregador do Forms. |

A aba `Base_HC` é utilizada para a lista de responsáveis do inventário. Caso o seu processo utilize estrutura diferente, revise a função de leitura correspondente antes da implantação.

### 4. Autorizar o acesso ao Drive

No menu personalizado **Painel Yubico** da planilha, execute **Autorizar acesso ao Drive**. Se solicitado, conclua a autorização com uma conta que tenha acesso à pasta raiz configurada. O sistema usa esse acesso para listar subpastas, carregar documentos e controlar os formulários de empréstimo.[1]

### 5. Publicar como Web App

Para disponibilizar o painel no navegador, use **Implantar → Nova implantação** e selecione o tipo **Aplicativo da Web**. Defina a execução e o público de acesso de acordo com as regras de segurança da organização. A implantação de um Web App é a forma oficial de disponibilizar uma interface HTML pelo Apps Script.[2]

Sempre que o código for alterado, acesse **Implantar → Gerenciar implantações → Editar**, escolha **Nova versão** e implante novamente. Caso essa etapa seja ignorada, o endereço publicado continuará executando a versão anterior.[3]

## Abas criadas ou utilizadas

| Aba | Origem | Descrição |
| --- | --- | --- |
| `Resumo_HC` | Existente | Base operacional principal de colaboradores, Yubicos e formulários. |
| `Base_HC` | Existente | Fonte para responsáveis disponíveis no inventário. |
| `Relatorio_Drive` | Criada/atualizada pelo sistema | Relatório consolidado dos arquivos encontrados no Drive. |
| `Inventario` | Criada/atualizada pelo sistema | Histórico dos inventários registrados. |
| `Emprestimos` | Criada automaticamente | Registro completo de empréstimos, formulários, status e devoluções. |
| `Informativos` | Criada automaticamente | Persistência dos indicadores gerais e por empregador. |

A aba `Emprestimos` é criada com as colunas abaixo caso ainda não exista.

| Ordem | Coluna |
| --- | --- |
| 1 | ID |
| 2 | CPF |
| 3 | Nome |
| 4 | Departamento |
| 5 | Função |
| 6 | Nível |
| 7 | Área |
| 8 | Turno |
| 9 | Setor |
| 10 | Empregador |
| 11 | Aparelho / Patrimônio |
| 12 | Data do empréstimo |
| 13 | Formulário |
| 14 | URL do formulário |
| 15 | Status |
| 16 | Data da devolução |

## Como operar o sistema

### Dashboard e atualização do Drive

O Dashboard apresenta os indicadores calculados a partir da base e o resumo por empregador. Use o botão **Atualizar informações do Drive** quando precisar reprocessar a estrutura de arquivos. O resultado da execução fica disponível no bloco de status e no registro de eventos.

O projeto também possui a função `criarGatilhoDiario()`, que cria um acionador diário para executar a atualização às 6 horas. Acionadores permitem a execução automática de funções em eventos ou horários definidos.[4]

### Envio de documentos

Na área de upload, pesquise e selecione a pasta de destino. O fluxo padrão de upload permite somente arquivos PDF. Após o envio, o sistema retorna o nome e o link do arquivo salvo no Drive.

### Inventário

Registre a quantidade disponível para cada modelo Yubico. O último registro é tratado como saldo atual e os registros anteriores ficam disponíveis no histórico. Os dados de inventário usam cache de propriedades para melhorar o carregamento da interface.

### Exportações da `Resumo_HC`

A aba **Exportações** suporta seleção múltipla para todos os filtros. Não selecionar valores em um filtro significa que aquele critério não restringirá o resultado.

| Filtro | Coluna da `Resumo_HC` | Regra |
| --- | --- | --- |
| Empregador | J | Lista os empregadores encontrados na base. |
| Departamento | B | Lista os departamentos encontrados na base. |
| Turno | H | Lista os turnos encontrados na base. |
| Possui Yubico | K | Campo preenchido = Sim; campo vazio = Não. |
| Possui formulário | L | Considera Sim ou Não. |
| Empregador do Forms confere | O | Considera Sim ou Não. |

Os resultados podem ser baixados em CSV ou XLSX. O arquivo Excel é gerado diretamente no backend, sem depender da conversão de uma planilha temporária.

### Empréstimos e devoluções

1. Abra **Empréstimos**.

1. Pesquise o colaborador por CPF ou nome e selecione o registro desejado.

1. Confirme os dados carregados da `Resumo_HC`.

1. Escolha o modelo: **Yubico 300**, **Yubico 325** ou **Yubico 600**.

1. Selecione uma pasta de destino para o formulário.

1. Anexe o formulário e registre o empréstimo.

Todo empréstimo nasce com o status **PENDENTE**. O histórico mostra a quantidade de dias pendentes, calculada a partir da data de empréstimo até a data atual. Ao registrar uma devolução, escolha entre manter o formulário no Drive ou movê-lo para a lixeira. A opção de exclusão envia o arquivo para a lixeira do Drive; ela não remove permanentemente o documento.

Use os filtros **Todos**, **Pendentes** e **Devolvidos**, além da busca textual, para localizar registros. O botão **Exportar CSV** gera um arquivo apenas com os empréstimos que estão visíveis de acordo com o filtro e a busca atuais.

### Informativos

Os indicadores gerais são preenchidos manualmente:

| Grupo | Campos |
| --- | --- |
| Gerais | HC atual, valor disponível para compra, estoque atual, aparelhos danificados, aparelhos com a Amazon e aparelhos com CNF5. |
| Por empregador | Quantidade pendente de devolução, sem saldo para desconto e desconto parcial. |

Os campos aceitam somente valores numéricos não negativos. A data de atualização exibida é calculada a partir da maior data válida salva na coluna **Atualizado em** da aba `Informativos`.

### Tela TV

O menu **Tela TV** abre a mesma implantação com o parâmetro `tela=tv`. Essa tela foi desenhada para acompanhamento em monitor ou TV e possui o slide de Informativos na segunda posição.

## Arquitetura e fluxo de dados

```
Resumo_HC ──────────────┬── Dashboard e indicadores
                        ├── Exportações filtradas
                        ├── Busca de colaborador para empréstimos
                        └── Empregadores dos Informativos

Google Drive ───────────┬── Upload de PDFs
                        ├── Árvore de pastas disponível na interface
                        └── Formulários de empréstimos

Inventario ─────────────── Histórico e saldo atual
Emprestimos ────────────── Status, devoluções e dias pendentes
Informativos ───────────── Indicadores manuais e por empregador
Relatorio_Drive ────────── Resultado da leitura do Drive
```

## Permissões e segurança

O projeto precisa de escopos de autorização para acessar planilhas e arquivos do Drive. A conta usada para executar o Web App deve ter permissão coerente com a política da organização e com a visibilidade da planilha e da pasta configuradas.[1]

| Recomendação | Motivo |
| --- | --- |
| Restrinja o acesso do Web App a usuários autorizados. | O painel manipula CPF, dados de colaboradores e links de formulários. |
| Use uma conta operacional com acesso controlado à planilha e à pasta raiz. | Evita falhas em uploads, leitura de pastas e ações de devolução. |
| Não compartilhe publicamente links de formulários. | Os links podem dar acesso a documentos internos. |
| Revise permissões periodicamente. | Mantém o princípio de menor privilégio. |
| Faça cópias de segurança antes de mudanças estruturais. | Protege o histórico operacional contra exclusões acidentais. |

## Validações locais disponíveis

Os testes são destinados ao ambiente de desenvolvimento e utilizam simulações dos serviços Google. Execute-os com Node.js antes de substituir os arquivos no Apps Script.

```bash
cd /caminho/para/o/projeto
python3 validar_sintaxe.py
node testar_servidor_exportacao.js
node testar_backend_emprestimos.js
node testar_backend_informativos.js
node testar_fluxo_filtros.js
node testar_exportacao_emprestimos.js
```

| Arquivo | Cobertura principal |
| --- | --- |
| `validar_sintaxe.py` | Extrai os scripts HTML e verifica a sintaxe com Node.js. |
| `testar_servidor_exportacao.js` | Valida filtros e geração de CSV/XLSX da `Resumo_HC`. |
| `testar_backend_emprestimos.js` | Valida consulta por CPF, cadastro, cálculo de dias, devolução e remoção opcional do formulário. |
| `testar_backend_informativos.js` | Valida a persistência de indicadores gerais e por empregador. |
| `testar_fluxo_filtros.js` | Valida a configuração e a renderização da interface de Exportações. |
| `testar_exportacao_emprestimos.js` | Valida o CSV de empréstimos com o filtro ativo. |

## Manutenção e evolução

Ao alterar os nomes de abas, colunas ou a estrutura da `Resumo_HC`, revise primeiro as constantes de configuração e as funções que usam índices de coluna. Alterações na regra de Yubico, formulário ou conferência devem ser atualizadas tanto na interface quanto no backend de exportação.

Para adicionar um novo modelo de Yubico, atualize simultaneamente a lista de opções do `Index.html` e a validação permitida em `registrarEmprestimo()` no `Code.gs`. Para criar novos indicadores manuais, inclua a chave na configuração de Informativos, adicione o campo à interface e mantenha a leitura e a gravação sincronizadas.

Antes de publicar uma alteração, faça a validação local, teste no editor com uma cópia de dados e, por fim, publique uma **nova versão** da implantação. Mantenha uma cópia versionada dos arquivos no Git para facilitar auditoria e reversão.

## Solução de problemas

| Sintoma | Verificação recomendada |
| --- | --- |
| Alterações não aparecem no painel publicado. | Publique uma nova versão em **Gerenciar implantações** e atualize a página sem cache. |
| Não é possível listar ou salvar arquivos no Drive. | Confirme `PASTA_ID`, permissões da conta executora e autorização do projeto. |
| Empréstimo não localiza o colaborador. | Verifique se o CPF está na coluna C da `Resumo_HC` e se os dados começam na segunda linha. |
| Filtros de Exportações não exibem opções. | Confirme a existência da `Resumo_HC`, os cabeçalhos e a publicação da versão atual do `Index.html`. |
| CSV de empréstimos vem vazio. | Revise o filtro de status e a busca textual aplicados no histórico antes de exportar. |
| A exclusão de formulário falha. | Confirme se a URL armazenada é de um arquivo do Drive acessível pela conta executora. |
| Informativos não atualizam a data. | Verifique se a aba `Informativos` possui datas válidas na coluna **Atualizado em**. |

## Referências

[1]: https://developers.google.com/apps-script/guides/services/authorization "Authorization for Google Services | Apps Script"

[2]: https://developers.google.com/apps-script/guides/web "Web Apps | Apps Script"

[3]: https://developers.google.com/apps-script/concepts/deployments "Create and manage deployments | Apps Script"

[4]: https://developers.google.com/apps-script/guides/triggers "Triggers | Apps Script"

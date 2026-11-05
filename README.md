# Automação de Datas: Evoluções × Planilha de Guias

Script em **Google Apps Script** que concilia as evoluções de um paciente (Google Docs) com as guias de atendimento registradas em uma planilha (Google Sheets). Ele corrige as datas das evoluções, marca na planilha as guias já utilizadas, reordena o documento por data e padroniza o layout, tudo a partir de um menu dentro da própria planilha.

> **English summary:** a Google Apps Script tool that reconciles clinical-note entries in a Google Doc against authorization records in a Google Sheet. It fixes entry dates, flags consumed records, reorders entries newest-first and normalizes page layout. It is idempotent, batch-optimized and designed to fail safely before touching the document.

---

## O problema

Em clínicas de terapia, cada atendimento gera uma evolução em um Google Doc e uma guia em uma planilha de faturamento. Na prática, as datas das evoluções ficam desencontradas das guias, o documento perde a ordem cronológica e o conferir tudo à mão consome horas por paciente. Este script automatiza esse trabalho e reduz o risco de erro humano.

## Impacto

Desenvolvido a partir de um processo manual real e adotado pela equipe de uma clínica de terapia multidisciplinar, onde está em uso ativo desde 2026:

- **Tempo por caso:** a conciliação de dados de faturamento levava cerca de **1h a 1h30** por caso e passou a ser feita em aproximadamente **20 minutos** em determinados casos.
- **Capacidade da equipe:** de cerca de **5–6** para cerca de **12–15 casos processados por dia**.
- **Performance:** um problema crítico de estouro de tempo de execução foi resolvido ao trocar milhares de chamadas individuais de API por um pequeno número de operações em lote.
- **Confiabilidade:** bugs de reconhecimento de padrões de texto (expressões regulares) foram identificados e corrigidos a partir de documentos reais, e a correspondência de registros por similaridade tolera erros de digitação.
- **Segurança operacional:** o modo de teste (dry run) permite validar o resultado sem risco antes de aplicar em produção, e as validações do script impedem o processamento incorreto.
- **Adoção:** a solução foi validada e adotada organicamente pela equipe, sem imposição.

## O que ele faz

- **Identifica o paciente automaticamente** pelo título do Doc, com tolerância a erros de digitação (similaridade de Levenshtein).
- **Pareia evoluções e guias por terapia**, em ordem cronológica, corrigindo a data de cada evolução.
- **Marca na planilha** as linhas já consumidas (cor `#d9ead3`), em uma única chamada em lote.
- **Sinaliza evoluções sem guia** com o marcador `TERAPIA A MAIS`.
- **Reordena o Doc por data**, da mais recente para a mais antiga, excluindo as evoluções sem guia.
- **Padroniza o layout:** cada evolução (exceto a primeira) começa em página nova, linhas em branco soltas e espaços indevidos são removidos.
- **Formata o documento** (Calibri, recuos, espaçamento e entrelinhas padronizados) sem alterar o conteúdo.
- **Modo de simulação (dry run)** que mostra o resumo sem gravar nada.

## Decisões de engenharia

| Tema | Abordagem |
|---|---|
| **Idempotência** | Linhas já pintadas são reconhecidas por um mapa de datas. Reexecutar no mesmo Doc não reescreve evoluções que já estavam corretas, mesmo quando uma guia nova entra no meio do histórico. |
| **Segurança antes da escrita** | Valida a consistência dos blocos e os tipos de elemento, copia todo o conteúdo e só então remove os originais. Qualquer falha anterior deixa o documento intacto. |
| **Referências, não índices** | Os blocos guardam a referência do elemento e a posição só é consultada no momento de mover, eliminando a aritmética de deslocamento de índices e a classe de bugs associada. |
| **Desempenho** | O corpo do Doc é lido uma vez e reaproveitado em memória. A planilha é lida e escrita em lote (valores, cores e `RangeList`). Remover as linhas em branco antes de copiar reduz muito o número de elementos movidos. |
| **Ordenação estável** | Datas iguais preservam a ordem original. |
| **Decisões ambíguas são recusadas** | Se dois pacientes têm nomes quase idênticos, o script não adivinha e pede ao usuário que corrija os nomes, evitando gravar no paciente errado. |
| **Confirmação em ações destrutivas** | Exclusões e remoções de elementos não suportados exigem confirmação explícita. |
| **Reprocessamento seguro** | O cabeçalho da evolução aceita tanto uma data quanto o marcador `TERAPIA A MAIS`, então um Doc já processado pode ser rodado de novo. |

## Pré-requisitos e formato esperado

### Planilha

Uma aba com os pacientes, e a **primeira linha** com as colunas (a busca ignora acentos e maiúsculas):

| Coluna | Conteúdo |
|---|---|
| `PACIENTE` | Nome do paciente |
| `TERAPIA` | Nome da terapia (por exemplo, `FONOTERAPIA`) |
| `DATA DE ATENDIMENTO` | Data da guia |

A planilha pode conter **todos os pacientes**. O script seleciona o paciente pelo título do Doc.

### Google Doc

- O **título do Doc** deve ser o nome do paciente, como está na planilha.
- Cada evolução começa com um cabeçalho no formato `Evolução dia DD/MM/AAAA das ...`.
- Logo depois vem a linha `Procedimentos: NOME DA TERAPIA`.

O regex do cabeçalho (`regexCabecalho`) e o de procedimento (`REGEX_PROCEDIMENTO`) são específicos desse fluxo. Para outro formato de documento, ajuste esses dois pontos.

## Instalação

1. Abra a planilha de guias e vá em **Extensões → Apps Script**.
2. Cole o conteúdo de `automacao_datas.gs` no editor e salve.
3. Recarregue a planilha. O menu **🔧 Automação de Datas** aparece na barra.
4. Na primeira execução, o Google pede autorização para acessar planilhas, documentos e a interface. O script roda **somente na conta de quem o instala**.

## Como usar

1. No menu **🔧 Automação de Datas**, execute **🧪 Rodar Teste (Dry Run)**.
2. Cole o link do Google Doc do paciente e confira o resumo (evoluções atualizadas, já corretas e sem guia).
3. Se estiver certo, execute **✅ Aplicar Automação** e cole o link novamente.
4. O script corrige as datas, pinta a planilha, reordena o Doc, padroniza as quebras de página e formata o documento.

### Ferramentas independentes

| Item do menu | O que faz |
|---|---|
| **🔄 Reordenar evoluções do Doc** | Reordena usando as datas que já estão no Doc, sem consultar a planilha. |
| **📄 Padronizar quebras de página** | Só ajusta quebras de página e linhas em branco, sem tocar em datas ou na ordem. |
| **📝 Formatar documento** | Só aplica a formatação padrão ao corpo do Doc. |

## Configuração

As opções ficam no objeto `CONFIG`, no topo do arquivo:

```javascript
const CONFIG = {
  COR_PINTADA: '#d9ead3',        // cor que marca guias já utilizadas
  SIMILARIDADE_MINIMA: 0.85,     // limiar para casar título do Doc com o nome na planilha
  MARCADOR_SEM_GUIA: 'TERAPIA A MAIS',
  COLUNAS: { PACIENTE: 'PACIENTE', TERAPIA: 'TERAPIA', DATA: 'DATA DE ATENDIMENTO' }
};
```

## ⚠️ Avisos importantes

- **Faça um teste em uma cópia do Doc antes de usar em produção.** O script **exclui** evoluções marcadas como `TERAPIA A MAIS` e move blocos de texto. Se algo sair errado, use o **Histórico de versões** do Google Docs para reverter.
- **Dados sensíveis (LGPD).** O script abre documentos e planilhas que contêm dados de saúde. Ele não envia nada para fora da sua conta do Google, mas quem o utiliza continua responsável pelo tratamento adequado desses dados. **Nunca publique no repositório** Docs, planilhas, prints ou logs com informações de pacientes.
- **A planilha é processada por inteiro, ignorando filtros.** Se a aba estiver filtrada em outro paciente, as linhas pintadas podem ficar ocultas. O resumo avisa quando isso acontece.

## Limitações conhecidas

- Depende do formato de cabeçalho e da linha `Procedimentos:` descritos acima.
- Evoluções dentro de tabelas ou caixas de texto não são suportadas e fazem a reordenação abortar sem alterar o documento.
- Em documentos muito grandes, a execução está sujeita ao limite de tempo do Apps Script.

## Tecnologias

Google Apps Script (JavaScript ES6), `DocumentApp`, `SpreadsheetApp`.

## Licença

Distribuído sob a licença MIT. Veja o arquivo `LICENSE`.

## Autor

Lucas Santana Silva · [GitHub](https://github.com/Lucas-SantanaS) · [LinkedIn](https://www.linkedin.com/in/lucas-santana-b19bb13a8/)
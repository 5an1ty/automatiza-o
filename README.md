# Cadastro PCM

Aplicativo desktop em português para importar cotações PDF, extrair itens, revisar
informações e gerar cópias do formulário oficial Excel e PDF via Microsoft Excel.
Nenhum documento é enviado à internet. Pesquisa automática e OCR não estão disponíveis.

## Instalar e iniciar no Windows

Requisitos: Python 3.11+ de 64 bits e Microsoft Excel desktop instalado/ativado.
Abra `iniciar.bat`. Ele cria `.venv`, instala as dependências e inicia o programa.
Alternativamente, na pasta do projeto:

```powershell
py -3 -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python -m cadastro_pcm.main
```

Não execute o aplicativo como administrador. Instalação requer acesso ao PyPI.
As preferências são gravadas em `%LOCALAPPDATA%\CadastroPCM\settings.json`;
logs com rotação ficam na mesma pasta. Não há nome fictício de requisitante:
informe seu nome na primeira execução. Logs podem conter caminhos corporativos.

## Modelo oficial: inspeção obrigatória

Para homologação, mantenha o modelo oficial, a cotação original e um PDF
preenchido de referência fora do checkout. Não adicione documentos corporativos,
relatórios detalhados nem cópias preenchidas ao GitHub.

O mapeamento abaixo foi conferido estruturalmente contra o modelo disponibilizado
na etapa Prompt 04. Isso não substitui a autorização humana de cada versão do
modelo nem a conferência visual da saída do Excel no Windows. Nenhum modelo
corporativo foi recriado; os testes publicados usam documentos sintéticos.

Em **Configurações**, selecione o modelo e a aba nacional ou importada. Use
**Inspecionar modelo** para salvar um relatório JSON de todas as abas, células,
fórmulas, proteção, mesclagens, listas e impressão. Confira o mapeamento abaixo
contra o arquivo e as instruções corporativas; confirme também assinaturas vazias.
Somente marque a autorização após a conferência. A autorização é vinculada ao
SHA-256 do modelo; qualquer alteração exige nova conferência.

| Campo | Célula preliminar |
|---|---|
| Requisitante | C8 |
| Data | E8 |
| Área | C9 |
| Nome | A12 |
| Descrição técnica e NCM fiscal (`NCM:`) | A14 |
| Classificação | A16 |
| Tipo (nunca escrito) | C16 |
| Grupo UM | E16 |
| Grupo de Itens | A18 |
| Finalidade | C18 |
| Código SAP | E18 |
| Valor unitário | A20 |
| Equipamento | C20 |
| Importado? | A22 |
| Órgão fiscalizador | C22 |
| NCM inferior (texto SIM) | E22 |

Se o mapeamento diferir, **não autorize**: é necessária adaptação do mapeamento
por um desenvolvedor após análise do modelo. O aplicativo recusa fórmulas em
campos de destino, células bloqueadas sob proteção ativa, âncoras incorretas,
listas incompatíveis, validações não suportadas e ausência de área de impressão.
Não remove nem contorna proteção. Solicite modelo autorizado quando houver bloqueio.
Listas literais, intervalos e nomes definidos simples são suportados; listas
calculadas/dinâmicas e validações personalizadas são bloqueadas conservadoramente.

## Fluxo de uso

1. Configure requisitante, data, área, saída, modelo, aba e padrões.
2. Selecione ou arraste vários PDFs; remova entradas antes de processar se necessário.
3. Clique **Processar**. Cada item reconhecido aparece separadamente.
4. Abra **Ver texto do PDF** para consultar documentos com layout não reconhecido;
   use **Adicionar item manual** quando necessário. Páginas digitalizadas geram aviso.
5. Edite por duplo clique. As origens aparecem nos rótulos e evidências completas nas
   dicas dos campos. Valores alterados são marcados como manuais. Quantidade e total
   são preservados no registro, mas somente o unitário é escrito em A20.
6. Para consulta externa, pesquise fora do programa em fontes autorizadas e use
   **Registrar pesquisa** para informar valor, URL e eventual estimativa. O programa
   pede confirmação antes de substituir; a aprovação do item é invalidada.
7. Revise dados e fontes e marque a aprovação. Nome e requisitante são obrigatórios; o preço unitário
   confirmado também, salvo autorização individual de preenchimento posterior no Excel. NCM ausente exige informar o código, confirmar que não se aplica ou
   autorizar explicitamente a aprovação com essa pendência para preenchimento posterior. Oito dígitos válidos não comprovam classificação fiscal correta.
8. Gere Excel ou Excel + PDF para o item selecionado. Cada operação cria uma pasta
   exclusiva `NOME_identificador/cadastro.xlsx`, `registro.json` e, quando solicitado,
   `cadastro.pdf`. O JSON registra dados e evidências. Não há sobrescrita de saídas.
9. Visualize o PDF; confira cortes, descrição longa, estrutura, Tipo e assinaturas.
   **Imprimir…** abre o visualizador do Windows; use Ctrl+P e confirme lá. Não é
   enviado nenhum trabalho automaticamente à impressora.

Alterações nas configurações valem para próximos itens processados; itens existentes
podem ser editados individualmente. Remover um documento da lista de entrada não
apaga itens já extraídos. A lista de revisão é mantida em memória: gere os registros
antes de fechar. Reprocessar pede confirmação antes de descartar a lista atual.

## Preservação e exportação

O openpyxl lê e inspeciona o modelo, sem salvá-lo. **No Windows a escrita é
feita pelo próprio Excel via pywin32**, sempre em uma cópia exclusiva. A validação
preliminar verifica endereços, âncoras, fórmulas, proteção e listas conhecidas.
Antes de salvar, o Excel verifica também as validações nativas dos destinos,
inclusive as não expostas pelo openpyxl. Falhas descartam a cópia de trabalho.
Não se remove a proteção. C16 (Tipo), assinaturas e células fora do mapeamento
não são escritos. Ainda se exige conferir as posições reais das assinaturas.

Sem Windows, a escrita OOXML existente é mantida para desenvolvimento e testes
sintéticos: ela altera os valores mapeados e acrescenta um estilo de data para E8 na cópia. Não é homologação do
modelo oficial nem substituto de Excel para gerar PDF. No Windows, a ausência de
Excel/pywin32 gera mensagem de incompatibilidade; não há fallback silencioso.

A exportação cria uma cópia temporária do Excel preenchido e abre essa cópia em
uma instância isolada do Excel, com macros e atualização de links desativadas.
Exporta somente a aba pertinente. O modelo original e o Excel já salvo não são
alterados pela preparação da impressão. Na cópia temporária, aplica A4 vertical,
uma página de largura e, por padrão, uma de altura, e limpa as quebras de página.
As margens existentes são mantidas. Antes de reduzir, mede largura e altura reais
pelo Excel e bloqueia escala inferior a 85%. Se a proteção impedir limpar quebras,
a proteção é mantida e o ajuste Fit To é usado para ignorar as quebras manuais;
o resultado ainda precisa passar por todas as verificações do PDF. Se falhar,
a causa e a restrição de quebras são informadas. A escala é uma barreira
conservadora; legibilidade, objetos e sobreposições ainda exigem aceitação visual.

Se não houver área de impressão, os limites preliminares são A1:E33 (nacional)
e A1:E34 (importado). Antes de utilizá-los, o programa rejeita conteúdo de células
ou mesclagens fora deles. Áreas existentes também são verificadas. Objetos,
imagens e assinaturas vazias não podem ser certificados apenas por esse teste:
confira o relatório, o modelo e o PDF antes de autorizar a versão real.

O PDF deve existir, conter o nome do item, a data brasileira e os textos das
células preenchidas/institucionais da aba, ter o número esperado de páginas e
não ter páginas sem texto. A exportação também confere A4 vertical, limites de
texto e fonte mínima de 5 pontos; não certifica imagens nem aceitação visual. PDFs reprovados ficam em diagnóstico local, fora do caminho de impressão; consulte a seção de diagnóstico PDF. Quando a exportação falha
após a geração do Excel, a mensagem indica onde o Excel e registro JSON ficaram.
Use **Exportar PDF do Excel salvo** para tentar novamente sem preencher outra
cópia nem reprocessar a cotação. A instância COM é encerrada em sucesso e erro; não se encerram processos Excel
de outras sessões do usuário.

A conferência visual de sobreposição, assinaturas e equivalência ao exemplo continua
obrigatória: ela não pode ser certificada sem os documentos reais. O programa não
redimensiona o formulário ou elimina campos para fazer o texto caber.

## Extração e extensão

`item_extractor.py` contém estratégias para tabelas com cabeçalhos semânticos e
blocos rotulados. Novos fornecedores podem adicionar estratégias em `STRATEGIES`.
O parser não usa totais como preço unitário, não deduz preço por divisão e não
inventa NCM. Layouts desconhecidos exigem revisão/entrada manual; não se promete
reconhecimento universal. Evidências contêm arquivo, página e trecho. PDFs são
processados em thread; a interface permanece responsiva com ações bloqueadas
para evitar alterações concorrentes durante o trabalho.

`research.py` registra sugestões manuais com fonte HTTPS, data e estimativa.
Nenhuma pesquisa online foi simulada ou executada. Uma futura integração deve
exigir autorização antes de transmitir dados corporativos.

## Testes

```powershell
.venv\Scripts\python -m pytest -q
```

No Linux, crie um venv e instale `requirements.txt`; pywin32 só é instalado no Windows.
A interface pode ser verificada sem monitor com `QT_QPA_PLATFORM=offscreen`.
Os testes cobrem PDFs sintéticos, múltiplos itens, dinheiro, NCM, evidências,
proteção, validações, integridade do pacote Excel, nomes, revisão gráfica e fluxo
PDF → extração → Excel sintético. Não substituem aceitação com arquivos reais.

Para validar no Windows: execute o aplicativo, processe a cotação real, confira os
itens (Gaxeta 16: quantidade 9, unitário 239,00 e total 2.151,00, conforme o pedido),
autorize o modelo após inspeção e gere PDF. Compare visualmente ao exemplo real,
confira apenas a aba pertinente, páginas, quebras e campos manuais vazios.
A extração da cotação real e a escrita OOXML em cópia do modelo foram verificadas
no Linux na etapa Prompt 04; a exportação real via Excel COM permanece pendente.

## Arquivos

- `cadastro_pcm/main.py`: inicialização e logs.
- `cadastro_pcm/interface/main_window.py`: interface, edição, configurações e workers.
- `cadastro_pcm/core/models.py`: dados, proveniência e validação.
- `pdf_reader.py`, `item_extractor.py`: leitura e estratégias de extração.
- `settings.py`, `research.py`, `file_manager.py`: preferências, fontes e saídas.
- `excel_writer.py`, `pdf_exporter.py`: inspeção, cópia e exportação COM.
- `tests/test_workflow.py`: testes automatizados.
- `requirements.txt`, `iniciar.bat`, `.gitignore`: instalação e execução.

Revise as licenças de PyMuPDF e Qt antes de distribuir o aplicativo na empresa.


## Atualização funcional — validação local no Windows

O tema escuro e a organização da interface foram preservados. O botão
**Diagnóstico**, no cabeçalho, executa verificações em uma thread:

- Versão de Python, Windows, pywin32 e acesso a uma instância isolada do Excel.
- Pasta de saída com escrita e limpeza de pasta temporária.
- Presença das quatro abas corporativas e situação dos campos mapeados na aba escolhida.
- Tipo vazio, hash autorizado e limites de impressão.
- Exportação real de PDF sintético pelo Excel, com uma página verificada.
- Se o modelo estiver autorizado e sem bloqueios, exportação temporária do modelo
  em cópia temporária **sem preenchimento**, para conferir paginação.

O diagnóstico não gera cadastros fictícios no modelo. O teste sintético é um livro
separado, criado em memória. Os PDFs de teste ficam em pasta temporária local e
são removidos. O relatório salvável contém estados, versão de Python e endereços
de células; não contém requisitante, caminhos locais, textos de cotação, conteúdo
de células ou credenciais. A inspeção detalhada do modelo é um relatório separado,
contém conteúdo corporativo e deve continuar somente no computador autorizado.
A homologação visual permanece pendente mesmo se as verificações passarem.

### Atualizar uma instalação existente

Se você baixou por Git, na pasta do projeto:

```powershell
git pull origin main
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m cadastro_pcm.main
```

Se não tem Git, baixe o ZIP da branch main no GitHub e extraia em outra pasta;
execute `iniciar.bat`. As configurações existentes em LOCALAPPDATA são mantidas.
Os comandos utilizam o Python do ambiente virtual; não exigem `python` global,
privilégios de administrador ou caminhos do desenvolvedor. Python 3.14.10 de
64 bits foi informado como funcional pelo usuário; esta etapa foi testada aqui
em Python 3.12 no Linux, sem comprovar a instalação das dependências em 3.14.

### Aceitação com os documentos oficiais (somente local)

1. Mantenha os documentos corporativos fora do checkout, por exemplo em uma pasta
   de documentos da empresa. Não envie os documentos nem relatórios detalhados ao GitHub.
2. Atualize dependências e abra o aplicativo. Confira requisitante, pasta de saída,
   política de data, padrões e aba nacional/importada em Configurações.
3. Selecione o modelo, salve a inspeção e confira o mapeamento, desbloqueios,
   listas, assinaturas, datas de aprovação, logotipo e controle de alterações.
4. Confira os limites preliminares de impressão antes de autorizar o hash.
5. Execute Diagnóstico, salve seu relatório local e corrija erros/pendências.
6. Importe `exemplo de solicitação.pdf` e confira se a **extração real** produziu
   Gaxeta 16, quantidade 9, unitário R$ 239,00, total R$ 2.151,00 e NCM 4016.93.00,
   conforme os valores de referência informados. O parser não contém esses valores.
7. Importe também uma cotação com vários produtos: confira separação dos itens,
   páginas/fontes, descrições quebradas e campos pendentes. Duplicatas são
   sinalizadas nas observações, nunca removidas nem unidas automaticamente.
8. Corrija os dados, confirme NCM ou não aplicabilidade e aprove. Tipo não é editável.
9. Gere Excel + PDF. Confira no Excel a preservação das quatro abas, validações,
   proteção, assinatura e Tipo vazios. Compare o PDF ao formulário preenchido de
   referência: não pode haver cortes, sobreposição, páginas extras ou conteúdo ausente.
10. Repita com a aba importada e com uma falha controlada (ex.: destino sem permissão).
    Confira a mensagem, o modelo original intacto e o encerramento da instância Excel.
11. Feche e reabra o aplicativo: preferências devem persistir; a tabela começa vazia,
    sem itens demonstrativos. Registros de revisão em memória não são restaurados.

### Resultado desta etapa

Na validação inicial, os documentos oficiais não estavam disponíveis. Na etapa
Prompt 04, os anexos foram inspecionados localmente, sem publicação. Exportação
real pelo Excel **não foi executada** no ambiente Linux. Testes automatizados utilizam somente PDFs e XLSX
sintéticos; casos COM usam objetos simulados explicitamente. Incluem colunas sem
bordas, descrições multilinha, duplicatas, múltiplas páginas, campos ausentes,
NCM, persistência, limites de impressão, falha de exportação, limpeza e proteção.
Os testes portáveis não iniciam Excel mesmo quando executados no Windows.
Não há OCR, consulta automática, transmissão externa ou homologação fiscal automática.


## Prompt 04 — correções do fluxo principal

- A extração reconhece os cabeçalhos `R$ Unit.` e `R$ Total`, mantendo a
  associação posicional das células da tabela, inclusive a coluna vazia de código
  e a coluna de IPI. Total e impostos nunca substituem o preço unitário.
- NCM rotulado no início, final ou entre parênteses é separado do nome/descrição.
  Um código inicial separado por hífen também é reconhecido. O formato armazenado
  e exibido é `XXXX.XX.XX`; a evidência conserva o trecho original. Códigos
  divergentes ficam pendentes. Números de contato e códigos de fabricante não
  são usados como fonte de NCM. Isso não valida a classificação fiscal.
- Quantidade permanece interna, no registro JSON e na verificação de coerência
  do total; não ocupa coluna, campo de edição ou célula no formulário.
- Ao importar, o primeiro PDF é selecionado quando não houver seleção. Remover
  e Ver texto seguem a seleção real. Remover uma entrada preserva itens extraídos,
  revisões e arquivos já gerados. Reprocessar substitui a lista somente após
  confirmação; exibe conclusão, ausência de itens ou causa da falha.
- A área Gerar documentos mostra os campos pendentes, a necessidade de aprovação
  ou o problema do modelo. O mesmo preflight verifica hash, aba, proteção, listas
  e impressão antes da habilitação e novamente antes de escrever. A verificação
  preliminar não comprova instalação do Excel nem permissão de escrita na saída;
  falhas nessas operações são apresentadas durante a geração e no Diagnóstico.
- A inspeção inclui listas estendidas `x14` do Excel. Campos opcionais já vazios
  no modelo permanecem intactos, em vez de receber uma nova atribuição vazia
  incompatível com uma lista. Valores preenchidos continuam sujeitos às listas,
  incluindo a validação nativa no Windows. Proteções não são removidas.
- No backend OOXML de testes, somente as células mapeadas são substituídas; o
  restante do XML, namespaces, extensões e componentes do arquivo são preservados.
  No Windows, continua sendo usada uma cópia editada pelo Microsoft Excel.
- Sucesso exige a existência dos arquivos. Uma falha na exportação indica a
  pasta do Excel e do registro preservados, sem anunciar um PDF inexistente.

### Validação desta correção

Com Python 3.12/Linux: **52 testes passaram**, incluindo os 34 casos anteriores
com ajustes para o NCM formatado, a coluna removida e o preflight mais rigoroso.
Os novos casos usam exclusivamente documentos sintéticos e exercitam também os
workers reais da interface, seleção, visualização de texto, reprocessamento,
edição, aprovação, bloqueios e gravação de Excel pelo botão de geração.

Os arquivos oficiais disponibilizados foram usados somente em validação local:
a cotação produziu Gaxeta 16, unitário 239,00, total 2.151,00 e NCM 4016.93.00.
O fluxo da interface gerou uma cópia OOXML de teste do modelo, preservando os
originais, proteção, sete listas estendidas da aba, Tipo e assinaturas vazios e
todos os componentes ZIP fora da aba preenchida. A aprovação nesse ensaio foi
apenas de teste; não configura autorização ou homologação corporativa.

O PDF preenchido fornecido é referência visual, não uma exportação produzida
pelo programa nesta máquina. A4, orientação vertical e uma página de largura
continuam sendo as opções de exportação solicitadas. **Excel COM e geração real
de PDF permanecem pendentes de execução no Windows**, assim como equivalência
visual, legibilidade, logotipo, assinaturas e quebras. Os testes COM publicados
usam simulações. Python 3.14 no Windows não foi executado neste ambiente.


## Prompt 05 — revisão simples, data e nova exportação

A revisão principal agora apresenta somente nome, descrição técnica, NCM e preço
unitário. Fabricante, código do fabricante, quantidade e total não aparecem no
editor, na tabela ou nos campos do painel de detalhes. As configurações e a aba
**Cadastro e solicitação** conservam requisitante, política de data, classificação
e os padrões administrativos. A pesquisa manual mantém descrição, NCM e preço;
nenhuma pesquisa automática foi adicionada. Campos antigos continuam legíveis no
JSON para compatibilidade, mas fabricante/código não são extraídos nem anexados
automaticamente ao formulário. Quantidade e total continuam apenas internos.

A identificação depende de cabeçalhos semânticos de tabelas ou campos de item
rotulados. Regras de contexto distinguem duração/prazo, pagamento, frete, dados
bancários, totais e seções comerciais. Essas seções encerram a associação com o
item anterior, evitando que preços de outra seção sejam herdados. Uma seção
comercial não é reaberta por uma descrição genérica; uma nova seção de itens ou
um produto/material/serviço explicitamente rotulado pode iniciar outro item.
Nomes legítimos de serviços, como transporte rodoviário, continuam aceitos.
Layout ambíguo permanece pendente para cadastro manual; não há promessa de
reconhecimento universal. Trechos comerciais removidos conservam evidência para
conferência. Itens legados cujo nome é uma condição comercial não podem ser gerados.

A14 recebe somente a descrição técnica e uma linha `NCM: XXXX.XX.XX`. NCM já
presente é normalizado sem duplicação; divergência com o campo fiscal exige
revisão. NCM confirmado como não aplicável não gera código fictício. Fabricante,
código de fabricante e condições comerciais não são adicionados à descrição.

E8 recebe uma data verdadeira e `NumberFormat = "dd/mm/yyyy"`, mantendo a política
atual ou fixa. Depois de salvar, a cópia é reaberta e a data/formatação são
verificadas. No backend OOXML de testes, acrescenta-se um estilo exclusivo para
E8 sem substituir estilos existentes. A data de emissão institucional não muda.

Se o PDF falhar, a interface conserva a referência ao Excel e mostra sua pasta,
a causa e o botão **Exportar PDF do Excel salvo**. A nova tentativa usa o mesmo
arquivo, aba e parâmetros registrados na geração, sem reler a cotação nem criar
outro cadastro. Não substitui PDFs existentes. Editar/reprocessar os itens invalida
a associação da saída anterior. A lista de itens e essas associações continuam
restritas à sessão atual do aplicativo.

### Verificações do Prompt 05

**91 testes passaram e nenhum falhou** em Python 3.12/Linux. Incluem condições
comerciais variadas, associação de preços, serviços legítimos, interface simples,
compatibilidade de registros, descrição final exata, datas nos sistemas Excel
1900/1904, preservação de estilos institucionais, políticas de data, falha e nova
exportação sem reprocessar. O fluxo de exportação nativa foi testado com COM
simulado: cópia temporária, remoção de quebras, A4, ajuste de largura/altura,
limite de escala, validação de conteúdo/data e PDF sintético de uma página,
incluindo o caso de uma folha protegida que recusa excluir quebras manuais.
Um caso adicional garante que substituir uma célula vazia no backend OOXML não
consuma o rótulo seguinte: todas as células fora do mapeamento são comparadas.

A cotação oficial disponibilizada foi novamente processada pelo worker e pela
revisão da interface: Gaxeta 16, R$ 239,00, NCM 4016.93.00. O botão Gerar Excel
produziu uma cópia OOXML local cuja descrição é exatamente `Gaxeta 16` seguida
de `NCM: 4016.93.00`. E8 foi conferida após reabrir, com data verdadeira e formato
brasileiro. Modelo, cotação e referência PDF permaneceram intactos; Tipo,
assinaturas, demais células institucionais e sete validações foram preservados.
Os documentos e saídas locais não integram o repositório.

**A exportação real pelo Microsoft Excel e a impressão ainda não foram executadas
neste ambiente Linux.** Não se declara homologação de uma página sem cortes,
legibilidade ou logotipo com base nas simulações. No Windows, atualize o projeto,
execute `iniciar.bat`, confira/autorize o modelo e gere Excel + PDF. Confirme a data,
formulário inteiro, assinaturas vazias, logotipo e controle de alterações. Se
falhar, use a nova exportação do Excel salvo e registre a causa exibida. Não se
remove proteção para tentar forçar uma exportação. Python 3.14/Windows permanece
pendente de execução nesta etapa.

## Datas COM e propostas com tabelas comerciais multilinha

A comparação literal do formato de E8 foi substituída por três verificações
separadas. No Windows, a data é escrita como serial numérico usando **Value2** e
`NumberFormat = "dd/mm/yyyy"`. Após salvar e fechar, o Excel reabre a cópia somente
para leitura: o valor é convertido pelo sistema de datas 1900/1904, o formato é
conferido semanticamente e a exibição é examinada à parte. Também se mantém a
verificação independente do arquivo salvo pelo leitor OOXML.

Formatos equivalentes como `dd/mm/yyyy;@`, prefixos de localidade e separadores
entre aspas/escapados são aceitos. O formato local `dd/mm/aaaa` é reconhecido.
Datas erradas, datas armazenadas como texto, formato americano, ano com dois
dígitos e formato General continuam bloqueados. `Value2` numérico não é um erro:
é a representação interna normal da data. `####` em Text indica largura
insuficiente, não uma data diferente; o PDF continua exigindo a data brasileira
visível, sem aceitar `####` ou o serial em seu lugar.

A validação anterior era excessivamente literal e não havia inspeção COM após
reabrir. Essas fragilidades foram corrigidas; a causa exata observada na máquina
do usuário ainda precisa ser confirmada pelo diagnóstico abaixo, pois este
ambiente Linux não possui Microsoft Excel real.

### Diagnóstico local de E8 no Windows

Na pasta do projeto atualizado, substitua o caminho pelo modelo oficial local:

```powershell
.\.venv\Scripts\python.exe -m cadastro_pcm.core.date_diagnostics --template "C:\Documentos\tela de cadastro.xlsx" --output-dir "$env:LOCALAPPDATA\CadastroPCM\diagnostico-data"
```

Sem `--date`, utiliza a data atual do computador. Para verificar uma data fixa,
repita acrescentando `--date 2026-11-05`. Para conferir também sua apresentação no
PDF, acrescente `--export-pdf`. A saída informa PASSOU/FALHOU e o caminho do JSON;
o processo retorna código 0 apenas quando o teste passa e o original está intacto.

O teste cria uma cópia exclusiva, altera somente E8, salva e reabre via Excel.
Preserva a proteção e recusa E8 bloqueada ou com fórmula. O relatório registra
somente a data esperada, Value2, sistema de datas, NumberFormat, NumberFormatLocal,
Text e resultados separados (`date_ok`, `format_ok`, `display`). A cópia permanece
para inspeção inclusive em falha; nenhum conteúdo de outras células é incluído
no relatório. Os arquivos são locais: não publique o modelo, a cópia nem os PDFs.
O teste não preenche nem aprova um cadastro e não altera a emissão institucional.

### Interpretação de tabelas

Cabeçalhos são combinados por coluna em até quatro linhas, reconhecendo, entre
outros, **Valor Unitário C/ Impostos** e **Valor Total C/ Impostos**. Não se exige
`R$` nas células: `1.422,00`, `R$ 1.422,00` e `1422,00` representam o mesmo valor
quando associados à coluna unitária. IPI/ICMS não são colunas monetárias e
percentuais não podem preencher o preço. Quantidade × unitário serve somente
como conferência auxiliar; o total nunca substitui a identificação do unitário.

A leitura compara tabelas com linhas e tabelas alinhadas por palavras. Quando o
detector genérico mistura texto de escopo com a tabela, há um fallback baseado
nas coordenadas das palavras: delimita o cabeçalho, alinha suas continuações e
associa dados às mesmas colunas. Detecções sobrepostas são selecionadas pela
clareza dos cabeçalhos/dados, sem duplicar itens. Uma caixa decorativa não impede
a leitura de outra tabela sem bordas na mesma página. Ambiguidade permanece
pendente, sem escolher arbitrariamente o maior valor encontrado.

Uma seção **Condições comerciais** pode conter uma tabela legítima de produtos:
o parser aceita sua estrutura e rejeita as linhas de prazo, pagamento, frete,
dados bancários ou totalização. Esses textos não se tornam produtos. Escopo
fornecido em seção própria só enriquece a descrição comercial quando há uma
referência explícita de item, única em ambos os lados e no mesmo PDF. A evidência
registra sua página. Referências repetidas ou ausentes não provocam associação
por suposição; descrições técnicas existentes não são substituídas silenciosamente.
A interface de quatro campos principais e o texto `NCM:` no formulário permanecem.

### Resultados desta etapa

**124 testes passaram; nenhum falhou** em Python 3.12/Linux. Incluem os casos
anteriores, formatos equivalentes de data, serial nos sistemas 1900/1904, data
errada com aparência correta, largura insuficiente, persistência/reabertura COM
simulada e diagnóstico que preserva original e cópia em falha. Há propostas
sintéticas de cinco páginas com tabelas com/sem bordas, cabeçalhos multilinha,
impostos e múltiplos produtos. O fluxo do worker, revisão/aprovação e botão Gerar
Excel foi exercitado com MANIPULO COM TRAVA, NCM 7326.90.90 e unitário 1.422,00.
Os valores são dados dos testes; não há respostas fixas no parser.

A cotação real da Indata disponível foi reprocessada: Gaxeta 16, NCM 4016.93.00,
unitário 239,00, quantidade interna 9 e total 2.151,00. A cópia Excel OOXML reabriu
com E8 correta e todas as células/estilos fora do mapeamento preservados.

Na entrega inicial desta correção, o PDF real de cinco páginas da Pharmainox
ainda não estava disponível; foram usados documentos sintéticos. A validação
posterior com o anexo real está descrita abaixo. **COM, Windows 11/Python
3.14 e exportação real pelo Excel não foram executados aqui.** Execute o diagnóstico
local e o fluxo real no Windows antes da aceitação final. Não se declara resolvida
a falha específica do computador nem homologada a impressão somente com simulações.


## Validação posterior com o PDF real da Pharmainox

Após o envio do PDF original de cinco páginas, a falha de preço foi reproduzida:
o detector de linhas criou colunas vazias adicionais, com os valores deslocados
em relação aos índices dos cabeçalhos. Os cabeçalhos também estão centralizados
verticalmente em alturas diferentes. A extração anterior encontrava o item e o
NCM, mas deixava o unitário e o total pendentes.

O fallback posicional agora inclui as linhas acima e abaixo do cabeçalho principal,
combina fragmentos de coluna por sobreposição horizontal e associa as palavras
dos dados pelos centros das colunas. Isso evita depender da borda esquerda de
um título centralizado ou de células artificiais do detector de linhas. Uma
extração com cabeçalho consistente e valores associados corretamente prevalece
sobre a grade fragmentada, sem gerar itens duplicados. Legendas numeradas de
imagens são excluídas da descrição técnica; os dados efetivos do escopo e suas
evidências são mantidos. Não foram adicionadas respostas fixas por fornecedor.

**127 testes passaram, nenhum falhou.** Novas regressões sintéticas reproduzem
cabeçalhos em alturas diferentes, texto centralizado, grade com colunas vazias
extras e legenda de imagem dentro do escopo. O PDF corporativo não integra os
testes publicados nem o repositório.

No PDF real, o fluxo de importação pela interface identificou apenas um item:
**MANIPULO COM TRAVA**, NCM **7326.90.90**, unitário **R$ 1.422,00**, quantidade
interna **8** e total **R$ 11.376,00**. A descrição foi associada ao escopo da página
3 pela referência explícita do item, sem incluir prazo de entrega, exclusões ou
legenda da imagem. A revisão e aprovação de teste pelo editor permitiram gerar
uma cópia Excel pelo botão da interface no backend OOXML do Linux. A cotação
real da Indata também foi revalidada: Gaxeta 16, NCM 4016.93.00 e unitário 239,00.

As cópias foram reabertas para conferir valores, E8 em data verdadeira e formato
brasileiro, Tipo/SAP/equipamento vazios, sete validações, proteção e todas as
células/estilos fora do mapeamento. Os dois PDFs e o modelo originais permaneceram
intactos. Aprovações e saídas foram apenas ensaios locais, sem homologação nem
configuração de autorização corporativa persistente.

**A gravação COM e o PDF final pelo Microsoft Excel continuam pendentes de teste
real no Windows.** A validação dos PDFs de entrada e da cópia OOXML não comprova
impressão, paginação ou legibilidade do formulário exportado pelo Excel.

## Correção pontual de apresentação regional da data E8

A falha «exibição da data diverge» vinha de comparar literalmente `Range.Text`
com `DD/MM/AAAA`, mesmo depois de validar o serial e o formato. Agora os três
resultados permanecem separados: **Value2** é a autoridade para o valor da data;
**NumberFormat** precisa continuar sendo brasileiro; **Text** é interpretado
segundo a ordem declarada por **NumberFormatLocal**. Também se observa **Value**,
que pode ser um variant de data, sem usá-lo para substituir o serial correto.

São reconhecidas apresentações numéricas regionais com separadores diferentes,
zeros opcionais, marcas direcionais e dígitos Unicode. Máscaras locais de dia/mês/
ano em português, inglês, francês e alemão são reconhecidas. Se a máscara local
não for interpretável, só se usa a ordem explicitamente declarada pelo formato
invariante; não se tenta adivinhar outra ordem para fazer a comparação passar.
Uma data invertida, valor divergente, formato inválido ou apresentação sem causa
confirmada continuam bloqueados. Não se aceita texto arbitrário como data real.

Antes de qualquer falha de validação COM de E8, grava-se um relatório técnico
local em `%LOCALAPPDATA%\CadastroPCM\diagnostico-data\diagnostico_data_*.json`.
A mensagem indica seu caminho. O relatório registra Value, Value2, Text,
NumberFormat, NumberFormatLocal, largura em pontos, largura em caracteres,
data esperada, ordem usada e os resultados individuais. Tentativas de ajuste
ficam no mesmo relatório. Ele contém somente observações de E8; nomes/caminhos
não são incorporados e textos não datáveis ou literais de formato desconhecidos
são omitidos. Se a verificação independente OOXML falhar, há um relatório próprio
que marca as propriedades COM não observadas como nulas. Não publique relatórios
nem cópias do formulário no GitHub.

Se aparecer `####`, a cópia pode receber ajuste **somente em E8**: barras literais,
localidade pt-BR e, havendo largura suficiente para leitura conservadora a oito
pontos, ShrinkToFit. Não se aplica AutoFit, não se alteram largura de colunas,
altura de linhas, mesclagens ou proteção. Se o ajuste for recusado ou a largura
for insuficiente, a operação falha com a causa e o relatório, sem modificar o
modelo original. Após salvar e reabrir, valor, formato e apresentação são
conferidos novamente.

Na cópia temporária de impressão, a apresentação exata é conferida e, quando
necessário, reaplicada usando a máscara local `dd"/"mm"/"aaaa` ou a máscara
invariante `[$-416]dd"/"mm"/"yyyy`. Uma representação regional equivalente pode ser aceita
na leitura do Excel, mas **o PDF exige o texto exato DD/MM/AAAA com barras e uma
data legível de pelo menos oito pontos**. A leitura do conteúdo do PDF, A4,
quantidade de páginas e preservação do formulário continuam sendo verificadas;
não se informa sucesso quando apenas o serial ou `####` estiverem visíveis.

**149 testes passaram, nenhum falhou**, em Python 3.12/Linux. Os novos casos
cobrem serial correto, diferença regional confirmada, ordem ambígua não inferida,
`####`, largura reparável e insuficiente, data realmente errada, formato inválido,
relatórios antes de exceção e sem dados pessoais, salvamento/reabertura, barras
literais no PDF, fonte mínima da data e dimensões do formulário preservadas.
Os casos de COM usam simulações explícitas, incluindo a configuração regional;
**a gravação e a exportação pelo Excel real ainda dependem do Windows do usuário**.
Use o diagnóstico local de E8 documentado acima, também com `--export-pdf`, para
confirmar a apresentação na instalação real antes da aceitação final.

## Recuperar o diagnóstico quando o JSON não aparece

A mensagem de E8 não exige mais localizar manualmente o arquivo automático.
Sua janela oferece **Salvar diagnóstico…**, **Copiar diagnóstico** e os detalhes
em JSON. Os dados sanitizados de E8 são mantidos em memória durante a operação,
inclusive quando o erro é encapsulado pelo worker ou ocorre na exportação PDF.
Salvar/cancelar/copiar não fecha a janela; use **Fechar** ao terminar.

A gravação automática agora é confirmada por releitura do JSON. Se a pasta padrão
não permitir gravar, tenta uma pasta temporária e informa o caminho confirmado.
Se nenhuma gravação for possível, a mensagem diz isso explicitamente, sem indicar
um arquivo inexistente, e os botões continuam disponíveis com os dados em memória.
Se um arquivo já confirmado desaparecer depois, o diagnóstico ainda pode ser
copiado ou salvo em outra pasta. Ao iniciar outra operação, o diagnóstico anterior
é descartado para não ser associado a um erro diferente.

Atualize o projeto, execute novamente a geração e, se E8 falhar, escolha
**Salvar diagnóstico…** para selecionar uma pasta ou **Copiar diagnóstico** para
obter o JSON sem depender de arquivos. A confirmação de salvamento mostra o
caminho do arquivo verificado. Os botões exportam somente o diagnóstico sanitizado
de E8, sem nomes, caminhos de usuário ou conteúdo de outras células. Não publique
os relatórios no GitHub.

**158 testes passaram, nenhum falhou** em Linux. Incluem falha de permissão,
arquivo removido antes/depois da confirmação, falha total de gravação, retenção
em memória, erros encapsulados pelo worker, erro de PDF com Excel preservado,
salvamento manual, cópia e continuidade da janela de erro. As validações de valor,
formato e apresentação de E8 permanecem ativas. Isso corrige o acesso ao diagnóstico;
não estabelece a causa específica da divergência que ocorre no Excel do usuário.
A reprodução e validação dessa divergência ainda exigem as observações reais do
Windows, obtidas pelo novo diálogo ou pelo diagnóstico local já documentado.


## Reaplicação do formato local após diagnóstico real de E8

O diagnóstico fornecido pelo usuário confirmou serial correto, data esperada e
formato brasileiro válidos, com 134,25 pontos de largura, mas Text não reconhecido.
O texto havia sido omitido pelo filtro de privacidade, portanto sua representação
exata ainda não estava disponível. A rotina anterior só tentava reparar largura
ou preparar o PDF: não reaplicava formato nesse caso de apresentação desconhecida.

Agora, na cópia de trabalho e **somente se valor e formato armazenados forem
válidos**, tenta primeiro `NumberFormatLocal = 'dd"/"mm"/"aaaa'`. Se não confirmar
a apresentação, tenta a máscara invariante pt-BR com barras literais. Cada tentativa
é inspecionada novamente: não se modifica Value2, não se aceita texto por suposição
e não se confunde uma data numérica realmente diferente com problema de formato.
A conferência após salvar e reabrir permanece. Larguras de colunas, estrutura,
proteção e modelo original continuam preservados.

Os relatórios passam a identificar `text_kind` e preservar sequências restritas
de máscara de data, como `06/11/yyyy`, `dd/11/yyyy` ou `06/mm/2026`. São exemplos
sintéticos, não uma afirmação sobre o texto ocultado no relatório anterior.
Textos pessoais ou arbitrários continuam omitidos. Se o ajuste local não for
confirmado, a geração permanece bloqueada e os detalhes continuam disponíveis
nos botões Salvar/Copiar diagnóstico.

**167 testes passaram, nenhum falhou** em Python 3.12/Linux. Novos testes cobrem
máscaras parcialmente literais, aplicação local, fallback invariante, recusa de
ambos os ajustes, serial incorreto que não pode ser reparado como apresentação e
salvamento/reabertura da cópia com a máscara local confirmada. Os testes COM são
simulados. A confirmação da apresentação específica no Windows do usuário continua
necessária; não se declara que o texto ocultado era uma máscara literal nem que
a instalação real do Excel já foi homologada.

## Geração em lote de cadastros

No Windows com Microsoft Excel instalado, revise cada item e marque sua aprovação
humana no editor. Em **Gerar documentos**, clique em **Gerar todos os aprovados (N)**.
O contador considera itens explicitamente aprovados, sem bloqueios estruturais
ou pendências que não tenham sido autorizadas para preenchimento posterior.
A configuração e a autorização do modelo continuam obrigatórias; a compatibilidade
de cada item com a aba e suas validações é conferida durante o processamento.

Para gerar parte da lista, selecione linhas com Ctrl/clique ou Shift/clique e use
**Gerar selecionados**. Confirme a quantidade antes de iniciar. Itens sem aprovação
ou com pendências não autorizadas são ignorados e aparecem no resumo com o motivo. A aprovação não
é concedida automaticamente. Os botões de geração individual continuam disponíveis.

A geração ocorre em segundo plano, sequencialmente, com progresso por material.
Cada item recebe Excel, PDF e registro JSON separados, usando as mesmas rotinas da
geração individual. O lote fica em `lote_AAAA-MM-DD_HHMM_identificador`, dentro da
saída configurada; cada material tem uma subpasta com nome seguro e identificador
exclusivo, inclusive quando nomes coincidem. Execuções anteriores não são sobrescritas.
O modelo original, proteções, Tipo e assinaturas seguem as regras existentes.

O resumo mostra sucesso, erro ou exclusão por pendência, motivo e pasta por item.
Use **Abrir pasta do lote** para conferir os documentos. **Tentar novamente os itens
com falha** gera somente esses itens em novas subpastas do mesmo lote, após nova
confirmação e validação. Os arquivos anteriores permanecem intactos. Se a exportação
PDF falhar, o Excel e o registro já salvos são preservados; também é possível selecionar
o item e usar **Exportar PDF do Excel salvo**, sem gerar outro Excel.

A interface bloqueia alterações enquanto trabalha e aguarda a conclusão antes de
fechar. Cada operação nativa inicializa e libera COM na thread responsável, abre
uma instância isolada do Excel e encerra somente essa instância. Não há paralelismo
entre documentos do lote. Confira visualmente os PDFs antes de imprimir.

Os testes de lote usam planilhas sintéticas e exportação PDF/COM simulada no Linux:
validam lotes unitários e múltiplos, pendências, seleção, nomes repetidos, falhas,
continuação, preservação, nova tentativa, progresso e liberação de recursos. Não
substituem a validação de um lote no Windows com Excel real nem a homologação visual.

Validação desta entrega: **177 testes passaram**, nenhum falhou, em Python 3.12/Linux
com Qt offscreen. A geração em lote via Excel real continua dependente de teste no Windows.

## Preço a preencher e aprovação de vários itens

Em **Editar e revisar item**, ao lado do preço, marque **Preencher preço posteriormente
no Excel** somente para os materiais em que deseja omitir o valor. A opção começa
desmarcada e é individual. O aviso informa que Valor Material ficará vazio no Excel
e no PDF; a tabela identifica esses materiais como **Preço a preencher**. Desmarque
a opção para informar o preço normalmente. Sem preço e sem autorização individual ou aprovação explícita com pendências, o
item continua bloqueado. Nenhum valor zero é criado para substituir um preço ausente.

A autorização fica no campo `price_fill_later` do registro JSON de cada material.
Registros antigos assumem `false`. Quando autorizada, a cópia gerada fica com A20
vazio, mesmo que o modelo tenha um preço anterior. Formatação e proteção permanecem;
uma fórmula ou um valor protegido que precise ser removido bloqueia a operação.
O modelo original nunca é alterado. A opção funciona na geração individual e em lote.

Na revisão, **Selecionar todos** seleciona os materiais visíveis; o mesmo botão passa
a **Desmarcar todos**. Ctrl/clique e Shift/clique continuam permitindo seleção parcial.
A seleção não altera dados nem aprovações. **Aprovar selecionados** apresenta os totais
de itens completos, preços a preencher autorizados e pendências, incluindo os motivos
por material. Só os itens aptos recebem aprovação após confirmação humana explícita;
preço e NCM ausentes podem ser autorizados no novo diálogo descrito abaixo.
A ação não concede autorização de omitir preço a nenhum item.

Para gerar dez materiais, revise os dados, autorize individualmente os preços que
ficarão em branco, selecione os itens e confirme **Aprovar selecionados**. Em seguida,
use **Gerar todos os aprovados** ou **Gerar selecionados**. Os materiais com bloqueios estruturais ou pendências não autorizadas
permanecem bloqueados. A revisão e a aprovação individual continuam
disponíveis; salvar uma revisão sem marcar a aprovação a revoga. Alterar um campo
depois de marcar a confirmação no editor também desmarca essa confirmação.

A lista de revisão continua em memória durante a sessão; a geração salva o registro
JSON com a decisão do preço. Esta atualização não adiciona restauração automática de
sessões. Os testes de PDF/COM usam simulações explícitas no Linux, não Excel real.

Validação desta atualização: **192 testes passaram**, nenhum falhou, em Python 3.12/Linux
com Qt offscreen. Inclui preço opcional, A20 vazio nos dois backends de escrita
(COM simulado), PDF com preço vazio (exportação nativa simulada), seleção, confirmação
de aprovação em lote, pendências, revogação após edição e os testes anteriores.

## Aprovar com pendências e revisar listas longas

Use **Aprovar selecionados** ou **Aprovar todos** (todos os materiais visíveis).
A nova janela tem resumo fixo, lista central rolável e rodapé fixo, sem mensagens
com centenas de linhas. Seu tamanho é limitado à área útil do monitor. A lista
mostra o nome, a situação de aprovação e as pendências por material; os botões
continuam acessíveis ao rolar 39, 100 ou mais registros.

Quando houver pendências, escolha explicitamente:

- **Sim, aprovar incompletos**: autoriza os campos preço e/ou NCM fiscal atualmente
  ausentes em cada item. Os materiais recebem **Aprovado com pendências** e podem
  ser gerados individualmente ou em lote. Isso não confirma os dados ausentes.
- **Não, somente completos**: aprova apenas os elegíveis sem pendências não
  autorizadas. Não concede novas autorizações aos incompletos.
- **Cancelar**: não altera os materiais.

A mesma janela aparece ao marcar a aprovação individual e salvar uma revisão
com campos pendentes. É possível salvar a revisão sem aprovação. Reabrir e editar
continua exigindo nova confirmação; a edição não reaproveita autorizações antigas.
A falta de nome/identificação ou de requisitante continua bloqueada. Fórmulas,
proteção, dados incompatíveis com o modelo e falhas de integridade continuam
sujeitos às validações existentes e não são dispensados pela aprovação humana.

O registro JSON contém `authorized_pending_fields`, com `unit_price` e/ou `ncm`
somente quando explicitamente autorizados. Registros antigos assumem lista vazia.
A autorização de um campo não dispensa outra pendência que surja depois. A opção
individual `price_fill_later` permanece disponível e compatível.

Preço ausente permanece vazio em A20, nunca zero; NCM fiscal ausente não adiciona
código à descrição nem marca automaticamente não aplicabilidade. A descrição
informada é preservada. O NCM inferior (E22, normalmente SIM) é um campo separado
e não é um código fiscal. Tipo, assinaturas, proteção e modelo original seguem
as regras existentes. Após preencher o Excel manualmente, exporte outro PDF: o
PDF já gerado não se atualiza automaticamente. Esse lembrete aparece na aprovação
e nos detalhes dos materiais com pendências autorizadas.

Validação desta entrega: **214 testes passaram**, nenhum falhou, em Python 3.12/Linux
com Qt offscreen. Inclui escolhas Sim/Não/Cancelar, aprovação individual, autorizações
por campo, proteção dos campos corporativos e geração em lote com preço/NCM ausentes.
As janelas de confirmação e resultado foram exercitadas com 1, 10, 39 e 100 itens,
rolagem e botões visíveis, incluindo uma área útil simulada de 640×480. A janela com
100 itens também foi inspecionada visualmente em modo offscreen. Os testes de
PDF/COM continuam usando simulações; não houve validação de Excel real nesta entrega.

## Exportação PDF: diagnóstico por etapa e recuperação

A falha atual do Windows precisa ser confirmada pelo relatório da instalação real.
Os testes Linux não estabelecem a causa de uma falha de COM, impressora/paginação
ou do modelo corporativo no computador do usuário.

A exportação agora registra separadamente preparação dos arquivos, abertura do
Excel, seleção da aba, apresentação da data, PageSetup, escala, chamada COM,
existência/tamanho do PDF, validação e encerramento. O relatório inclui horário UTC,
aba, área de impressão, parâmetros numéricos de PageSetup, páginas produzidas,
códigos de falha e HRESULT/SCODE disponíveis. Não inclui nomes de materiais,
conteúdo de células, caminhos pessoais nem mensagens COM brutas. O diagnóstico
sanitizado de E8 é incluído quando a data impede prosseguir.

Na cópia temporária, a seleção da aba substitui eventual agrupamento de abas e a
planilha é calculada antes de preparar a impressão. A exportação continua sendo
nativa, pela planilha via `ExportAsFixedFormat`, com caminho absoluto e área de
impressão respeitada. A4, margens existentes e a barreira de legibilidade de 85%
continuam válidos. Não se remove proteção nem se ajustam globalmente as dimensões
do formulário; o modelo e o Excel salvo permanecem intactos. A seleção explícita
previne exportação acidental de abas agrupadas, mas não comprova que essa tenha
sido a causa das duas páginas observadas anteriormente.

**PDF gerado, mas não aprovado na validação** significa que houve arquivo, mas
paginação, conteúdo, data ou layout não passaram. O arquivo é preservado como
`diagnostico_pdf/reprovado_<id>.pdf` na pasta da saída, acompanhado de relatório
JSON. Esses PDFs podem conter documentos corporativos: mantenha-os locais e não
os publique. Eles não aparecem como PDF aprovado na visualização/impressão do
aplicativo. Um marcador `.pdf.pendente` também bloqueia a impressão caso o Windows
impeça mover o arquivo reprovado. Falhas antes de criar o PDF e falhas ao encerrar
o Excel recebem mensagens distintas. O relatório permanece disponível em memória
para Salvar/Copiar mesmo se a gravação automática falhar.

Para verificar a instalação sem documentos empresariais:

1. Abra **Diagnóstico → Testar exportação PDF** no Windows.
2. O aplicativo cria um Excel sintético, exporta com o Excel instalado, valida uma
   página e remove os arquivos temporários tanto em sucesso quanto em erro.
3. Use **Copiar diagnóstico** ou **Salvar relatório…**. Somente o relatório
   sanitizado escolhido pelo usuário é preservado fora da pasta temporária.
4. Se o teste passar, tente exportar o cadastro real e compare a etapa da falha.
   O teste sintético não homologa o layout do formulário corporativo.

Para recuperar arquivos de sessões anteriores, use **Gerar PDF a partir de Excel
existente**. Selecione `cadastro.xlsx` junto do seu `registro.json`, confirme a aba
validada e escolha um novo destino PDF. A identificação, Tipo vazio, data e área de
impressão são conferidos; a exportação e a validação final são repetidas. O hash da
versão escolhida impede exportar silenciosamente uma versão alterada após a
conferência. Alterações manuais anteriores à seleção são uma nova versão a ser
conferida pelo usuário; não exigem reimportar a cotação ou aprovar materiais outra
vez. O registro local associa o arquivo ao fluxo do aplicativo, sem constituir
certificação de autenticidade ou homologação visual. Não se sobrescreve um PDF
existente: escolha outro nome ou mantenha a saída anterior fora do novo destino.

No resumo do lote, **Tentar novamente somente PDFs** reutiliza os Excel salvos dos
itens com erro, sem preencher novamente nem modificar cadastros bem-sucedidos.
**Salvar diagnósticos do lote…** exporta os relatórios sanitizados disponíveis.
As falhas individuais continuam permitindo os demais itens; se a própria chamada
`Quit` falhar, o lote é interrompido explicitamente para evitar acumular instâncias
Excel cuja liberação não pôde ser confirmada. Close, Quit e CoUninitialize são
sempre tentados; uma falha de limpeza não esconde o erro original da operação.

Depois de editar um Excel manualmente, gere outro PDF e confira cortes, logotipo,
bordas, assinaturas, Tipo, data e legibilidade antes de imprimir. O PDF anterior
não se atualiza automaticamente. Não foram introduzidas impressoras PDF virtuais,
LibreOffice ou serviços externos.

Validação desta entrega: **232 testes passaram**, nenhum falhou, em Python 3.12/Linux
com Qt offscreen. As simulações incluem exportação de aba específica, caminhos com
espaços, PageSetup, uma/duas páginas, PDF vazio ou ausente, conteúdo ausente, HRESULT,
permissões, destino existente/bloqueado, quarentena impedida, limpeza COM e
preservação de erro primário, reexportação de versão manual, teste sintético,
relatórios sem conteúdo corporativo e nova tentativa dos PDFs do lote.
**Microsoft Excel real não está disponível neste ambiente.** A causa atual no
Windows e a paginação/aceitação visual do modelo real permanecem pendentes do teste
local e do relatório copiado pelo usuário; não se declara correção definitiva com
base nestes testes simulados.

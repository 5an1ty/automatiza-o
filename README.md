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

Os três documentos de referência **não foram fornecidos** nesta sessão:

- `tela de cadastro.xlsx`
- `exemplo de solicitação.pdf`
- `ADESIVO TRAVA ROLAMENTO LOCTITE 648 50 ML HENKEL(1).pdf`

Não foi possível confirmar o mapeamento nem a aparência real. Nenhum modelo
corporativo foi recriado. Testes usam documentos sintéticos, exclusivamente.

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
| Descrição técnica, fabricante, código e NCM fiscal | A14 |
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
7. Revise dados e fontes e marque a aprovação. Nome, requisitante e preço unitário
   confirmado são obrigatórios. NCM ausente bloqueia a aprovação até que o código seja informado ou o
   responsável marque explicitamente que ele não se aplica no editor. Oito dígitos válidos não comprovam classificação fiscal correta.
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
sintéticos: ela altera apenas os valores mapeados na cópia. Não é homologação do
modelo oficial nem substituto de Excel para gerar PDF. No Windows, a ausência de
Excel/pywin32 gera mensagem de incompatibilidade; não há fallback silencioso.

A exportação abre uma instância isolada do Excel, desativa macros e atualização
de links, abre a cópia somente para leitura e exporta apenas a aba pertinente.
Aplica A4, orientação vertical e largura de uma página, sem salvar a mudança no
arquivo. A altura não é comprimida automaticamente: se houver mais páginas que
o esperado, a exportação falha para exigir revisão, evitando reduzir texto até
ficar ilegível. Quebras e legibilidade precisam de aceitação visual no Windows.

Se não houver área de impressão, os limites preliminares são A1:E33 (nacional)
e A1:E34 (importado). Antes de utilizá-los, o programa rejeita conteúdo de células
ou mesclagens fora deles. Áreas existentes também são verificadas. Objetos,
imagens e assinaturas vazias não podem ser certificados apenas por esse teste:
confira o relatório, o modelo e o PDF antes de autorizar a versão real.

O PDF deve existir, conter o nome do item, ter o número esperado de páginas e não
ter páginas sem texto. PDFs inválidos são removidos. Quando a exportação falha
após a geração do Excel, a mensagem indica onde o Excel e registro JSON ficaram.
A instância COM é encerrada em sucesso e erro; não se encerram processos Excel
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
Este teste real ainda não foi realizado; Linux não oferece Excel COM.

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
  original **somente para leitura, sem preenchimento**, para conferir paginação.

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

Documentos oficiais não estavam disponíveis. Exportação real pelo Excel **não foi
executada** no ambiente Linux. Testes automatizados utilizam somente PDFs e XLSX
sintéticos; casos COM usam objetos simulados explicitamente. Incluem colunas sem
bordas, descrições multilinha, duplicatas, múltiplas páginas, campos ausentes,
NCM, persistência, limites de impressão, falha de exportação, limpeza e proteção.
Os testes portáveis não iniciam Excel mesmo quando executados no Windows.
Não há OCR, consulta automática, transmissão externa ou homologação fiscal automática.

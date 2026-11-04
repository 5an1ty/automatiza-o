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
   confirmado são obrigatórios. NCM ausente fica pendente; a revisão humana decide
   sua aplicabilidade. Oito dígitos válidos não comprovam classificação fiscal correta.
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

O openpyxl lê e inspeciona o modelo. A gravação altera somente os valores das
células mapeadas no XML da aba de uma cópia; todos os demais componentes do pacote
XLSX permanecem byte a byte iguais, incluindo mídias e outras abas. O original é
somente lido. O estilo monetário e de data deve estar correto no modelo autorizado.
Tipo não é escrito e deve estar vazio; assinaturas permanecem intocadas. Exige-se
modelo original limpo: as posições de assinatura ainda não foram verificadas.

A exportação usa instância isolada do Excel, macros desabilitadas, sem atualizar
links, abre a cópia somente para leitura e exporta exclusivamente a aba escolhida,
respeitando sua área de impressão. Fecha livro e Excel ao terminar. Confere páginas
esperadas e nome do item no texto do PDF. Falhas não são declaradas como sucesso;
o PDF inválido é removido e o Excel já gerado permanece na pasta de saída.

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

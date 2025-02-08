# Dicionário Completo de Palavras-Chave, Objetos, Propriedades e Métodos em VBA

Este dicionário é uma referência para consultar os termos comuns na linguagem VBA, com seus significados e usos explicados de forma simples.

---

## Controle de Fluxo (Flow Control)

- **If...Then...Else** (Se...Então...Senão): Executa um bloco de código com base em uma condição.

  ```vba
  If x > 10 Then
      MsgBox "Maior que 10"
  Else
      MsgBox "Menor ou igual a 10"
  End If
  ```

- **Select Case** (Selecionar Caso): Alternativa para múltiplas condições.

  ```vba
  Select Case x
      Case 1
          MsgBox "Um"
      Case 2
          MsgBox "Dois"
      Case Else
          MsgBox "Outro valor"
  End Select
  ```

- **For...Next** (Para...Próximo): Laço de repetição com contador.

  ```vba
  For i = 1 To 10
      MsgBox i
  Next 
  ```

- **For Each...Next** (Para Cada...Próximo): Iteração para percorrer coleções.

  ```vba
  Dim celula As Range
  For Each celula In Range("A1:A10")
      celula.Value = "Teste"
  Next celula
  ```

- **Do While...Loop** (Faça Enquanto...Loop): Repetição enquanto a condição for verdadeira.

  ```vba
  Do While x < 10
      x = x + 1
  Loop
  ```

- **Do Until...Loop** (Faça Até...Loop): Repetição até que a condição seja verdadeira.

  ```vba
  Do Until x > 10
      x = x + 1
  Loop
  ```

- **While...Wend** (Enquanto...Fim Enquanto): Estrutura de repetição semelhante a `Do While`.

---

## Declaração de Variáveis (Variable Declaration)

- **Dim** (Dimensionar): Declara variáveis.

  ```vba
  Dim x As Integer
  ```

- **Static** (Estático): Mantém o valor da variável entre execuções.

  ```vba
  Static contador As Integer
  ```

- **Const** (Constante): Declara uma constante.

  ```vba
  Const PI As Double = 3.14159
  ```

- **Public** (Público): Torna a variável acessível em todos os módulos.

  ```vba
  Public valorGlobal As String
  ```

- **Private** (Privado): Limita o escopo ao módulo atual.

  ```vba
  Private nomeUsuario As String
  ```

- **ReDim** (Redimensionar): Redimensiona arrays.

  ```vba
  ReDim meuArray(10)
  ```

- **Option Explicit** (Opção Explícita): Exige declaração explícita de variáveis.

---

## Arrays (Matrizes/Vetores)

- **Array** (Matriz/Vetor): Coleção indexada de elementos do mesmo tipo.

  ```vba
  Dim meuArray(1 To 5) As String
  ```

- **LBound** (Limite Inferior): Retorna o índice inferior de um array.

  ```vba
  lowerBound = LBound(meuArray)
  ```

- **UBound** (Limite Superior): Retorna o índice superior de um array.

  ```vba
  upperBound = UBound(meuArray)
  ```

---

## Tipos de Dados (Data Types)

- **Integer** (Inteiro): Números inteiros (-32.768 a 32.767)
- **Long** (Longo): Números inteiros maiores
- **Double** (Duplo): Números com ponto flutuante
- **String** (Texto): Texto
- **Boolean** (Booleano): Valores `True` ou `False`
- **Variant** (Variante): Tipo genérico
- **Date** (Data): Datas e horas

---

## Objetos e Propriedades do Excel VBA (Excel VBA Objects and Properties)

- **Application** (Aplicativo): Representa o aplicativo Excel
- **ActiveCell** (Célula Ativa): Célula atualmente selecionada
- **ActiveSheet** (Planilha Ativa): Planilha ativa
- **ActiveWorkbook** (Pasta de Trabalho Ativa): Pasta de trabalho ativa
- **Cells** (Células): Todas as células em uma planilha
- **Columns** (Colunas): Referencia colunas
- **Rows** (Linhas): Referencia linhas
- **Range** (Intervalo): Intervalo específico de células
- **Sheets** (Planilhas): Todas as planilhas
- **Worksheets** (Planilhas de Trabalho): Apenas planilhas, excluindo gráficos
- **Selection** (Seleção): Item atualmente selecionado
- **ThisWorkbook** (Esta Pasta de Trabalho): Referencia a pasta de trabalho do código VBA
- **FileSystemObject** (Objeto do Sistema de Arquivos): Manipula arquivos e diretórios
- **Dictionary** (Dicionário): Estrutura de pares chave-valor
- **UserForm** (Formulário do Usuário): Cria formulários personalizados

---

## Métodos Comuns (Common Methods)

- **Activate** (Ativar): Ativa um objeto
- **Copy** (Copiar): Copia células, planilhas ou objetos
- **Delete** (Excluir): Exclui um objeto
- **Select** (Selecionar): Seleciona um objeto
- **ClearContents** (Limpar Conteúdo): Limpa o conteúdo das células
- **Close** (Fechar): Fecha uma pasta de trabalho
- **Open** (Abrir): Abre uma pasta de trabalho
- **Save** (Salvar): Salva a pasta de trabalho
- **SaveAs** (Salvar Como): Salva com um nome específico
- **Find** (Localizar): Procura um valor
- **Replace** (Substituir): Substitui valores
- **Add** (Adicionar): Adiciona elementos a coleções
- **Insert** (Inserir): Insere linhas ou colunas
- **Sort** (Classificar): Ordena dados em um intervalo
- **Show** (Mostrar): Exibe formulários personalizados

---

## Propriedades Comuns (Common Properties)

- **Value** (Valor): Obtém ou define o valor de uma célula
- **Text** (Texto): Retorna o texto exibido
- **Formula** (Fórmula): Obtém ou define uma fórmula
- **Address** (Endereço): Retorna o endereço de uma célula
- **Count** (Contagem): Retorna o número de itens
- **Offset** (Deslocamento): Referencia uma célula deslocada
- **Interior.Color** (Cor Interior): Define a cor de fundo
- **Font.Bold** (Negrito): Define se o texto é negrito
- **EntireRow** (Linha Inteira): Referencia a linha inteira
- **EntireColumn** (Coluna Inteira): Referencia a coluna inteira
- **Orientation** (Orientação): Define a direção da ordenação

---

## Funções de Conversão (Conversion Functions)

- **CStr** (Converter para String): Converte para String
- **CInt** (Converter para Inteiro): Converte para Integer
- **CLng** (Converter para Longo): Converte para Long
- **CDbl** (Converter para Duplo): Converte para Double
- **CBool** (Converter para Booleano): Converte para Boolean
- **CDate** (Converter para Data): Converte para Date

---

## Funções de Data e Hora (Date and Time Functions)

- **Now** (Agora): Retorna a data e hora atual
- **Date** (Data): Retorna apenas a data atual
- **Time** (Hora): Retorna apenas a hora atual
- **DateAdd** (Adicionar Data): Adiciona ou subtrai um intervalo de tempo
- **DateDiff** (Diferença de Data): Calcula a diferença entre datas
- **Year** (Ano), **Month** (Mês), **Day** (Dia): Extraem componentes de uma data

---

## Manipulação de Arquivos (File Manipulation)

- **Dir** (Diretório): Retorna nomes de arquivos que correspondem a um padrão
- **Kill** (Matar): Exclui um arquivo
- **Name** (Nome): Renomeia um arquivo
- **MkDir** (Criar Diretório): Cria um diretório
- **RmDir** (Remover Diretório): Remove um diretório

---

## Funções de String (String Functions)

- **Left** (Esquerda): Extrai caracteres da esquerda
- **Right** (Direita): Extrai caracteres da direita
- **Mid** (Meio): Extrai partes de uma string
- **Trim** (Aparar): Remove espaços em branco
- **LTrim** (Aparar Esquerda): Remove espaços à esquerda
- **RTrim** (Aparar Direita): Remove espaços à direita
- **Replace** (Substituir): Substitui ocorrências em uma string
- **Split** (Dividir): Divide uma string em um array
- **Join** (Juntar): Une elementos de um array em uma string
- **Len** (Comprimento): Retorna o comprimento de uma string
- **InStr** (Posição na String): Retorna a posição de uma substring

---

## Formatação de Células (Cell Formatting)

- **NumberFormat** (Formato Numérico): Define o formato numérico
- **Font.Size** (Tamanho da Fonte): Define o tamanho da fonte
- **Font.Color** (Cor da Fonte): Define a cor da fonte
- **Font.Italic** (Itálico): Define texto em itálico
- **Font.Underline** (Sublinhado): Define sublinhado
- **Borders** (Bordas): Configura bordas de células
- **WrapText** (Quebrar Texto): Habilita quebra de texto
- **MergeCells** (Mesclar Células): Mescla células

---

## Tratamento de Erros (Error Handling)

- **On Error GoTo** (Em Erro Ir Para): Direciona para um manipulador de erros
- **On Error Resume Next** (Em Erro Continuar Próximo): Ignora erros e continua executando
- **Err.Number** (Número do Erro): Retorna o código do erro

---

## Eventos Comuns (Common Events)

- **Workbook_Open** (Abrir Pasta de Trabalho): Executa ao abrir uma pasta de trabalho
- **Worksheet_Change** (Alteração na Planilha): Executa ao alterar uma célula
- **Worksheet_Activate** (Ativar Planilha): Executa ao ativar uma planilha
- **BeforeSave** (Antes de Salvar): Executa antes de salvar
- **BeforeClose** (Antes de Fechar): Executa antes de fechar
- **Worksheet_SelectionChange** (Mudança de Seleção): Executa ao mudar a seleção de células

---

## Constantes de Mensagens (Message Constants)

- **vbOKOnly** (Apenas OK): Apenas botão OK
- **vbOKCancel** (OK Cancelar): Botões OK e Cancelar
- **vbYesNo** (Sim Não): Botões Sim e Não
- **vbYesNoCancel** (Sim Não Cancelar): Botões Sim, Não e Cancelar
- **vbInformation** (Informação): Ícone de informação
- **vbQuestion** (Pergunta): Ícone de pergunta
- **vbCritical** (Crítico): Ícone de erro
- **vbExclamation** (Exclamação): Ícone de aviso

---

## Depuração (Debugging)

- **Debug.Print** (Imprimir Depuração): Imprime no imediato (immediate window)
- **Stop** (Parar): Pausa a execução
- **Break** (Quebrar): Define um ponto de interrupção no código
- **Watch** (Observar): Monitora valores de variáveis

---

## Configurações de Aplicativo (Application Settings)

- **Application.DisplayAlerts** (Exibir Alertas): Controla exibição de alertas
- **Application.StatusBar** (Barra de Status): Define mensagem na barra de status
- **Application.Cursor** (Cursor): Define o cursor
- **Application.Wait** (Aguardar): Pausa a execução por um tempo
- **Application.OnTime** (Agendar): Agenda execução de macro

---

## Constantes Específicas do Excel VBA (Excel VBA Specific Constants)

- **xlAscending** (Crescente): Ordenação em ordem crescente
- **xlDescending** (Decrescente): Ordenação em ordem decrescente
- **xlGuess** (Adivinhar): O Excel tenta determinar automaticamente se há cabeçalhos
- **xlFormulas** (Fórmulas): Pesquisa nas fórmulas das células
- **xlValues** (Valores): Pesquisa nos valores das células
- **xlWhole** (Inteiro): Corresponde ao conteúdo exato da célula
- **xlPart** (Parte): Encontra uma correspondência parcial do valor
- **xlByRows** (Por Linhas): Pesquisa por linhas
- **xlByColumns** (Por Colunas): Pesquisa por colunas
- **xlNext** (Próximo): Busca pela próxima ocorrência
- **xlPrevious** (Anterior): Busca pela ocorrência anterior
- **xlTopToBottom** (De Cima para Baixo): Ordenação de cima para baixo
- **xlLeftToRight** (Da Esquerda para Direita): Ordenação da esquerda para a direita
- **xlNormal** (Normal): Salva no formato padrão do Excel
- **xlEdgeBottom** (Borda Inferior): Refere-se à borda inferior
- **xlEdgeTop** (Borda Superior): Refere-se à borda superior
- **xlEdgeLeft** (Borda Esquerda): Refere-se à borda esquerda
- **xlEdgeRight** (Borda Direita): Refere-se à borda direita

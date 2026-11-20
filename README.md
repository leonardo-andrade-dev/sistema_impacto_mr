# Epimed Cultura

Sistema para converter os relatórios do Epimed em Word e cruzá-los com a
planilha de Cultura.

## Como usar

1. Inicie o sistema de uma das seguintes formas:
   - Execute `python scr\main.py` a partir da pasta do projeto; ou
   - Abra `dist\EpimedCultura.exe`, depois que o executável tiver sido gerado.
2. Clique em **Criar Pastas**. O sistema limpa o conteúdo anterior das pastas
   de trabalho, recria a estrutura e abre a pasta do projeto.
3. Confira o formato da planilha de Cultura. Se ela estiver em `.xls`, converta-a
   manualmente para `.xlsx` no Excel ou LibreOffice e salve como **Pasta de
   Trabalho do Excel (.xlsx)**. Se já estiver em `.xlsx`, não é necessária
   conversão.
4. Copie os arquivos `.docx` do Epimed para `Dados_Epimed`.
5. Copie a planilha de Cultura já convertida para `.xlsx` para `Culturas`.
6. No sistema, clique em **Realizar Conversão**. Os arquivos Word serão
   convertidos para planilhas `.xlsx` dentro de `Dados_Epimed\xlsx`.
7. Clique em **Gerar Base Final** para cruzar as informações.
8. A planilha resultante será salva em `Saida\Epimed_Cultura.xlsx`.

## Estrutura de pastas

```text
Culturas\             Planilha de Cultura no formato .xlsx
Dados_Epimed\         Relatórios do Epimed no formato .docx
Dados_Epimed\xlsx\    Arquivos convertidos pelo sistema
Saida\                Base final gerada pelo sistema
```

## Atenção

- Clique em **Criar Pastas** antes de começar cada processamento. Essa ação
  apaga os arquivos existentes nas pastas `Culturas`, `Dados_Epimed` e `Saida`
  (incluindo suas subpastas) antes de recriá-las. Faça cópias de segurança dos
  arquivos que deseja guardar.
- Atualmente, o sistema aceita a planilha de Cultura somente em `.xlsx`.
  Arquivos `.xls` precisam ser convertidos manualmente antes de serem copiados
  para `Culturas`.
- Ao usar o executável, as pastas de trabalho serão criadas ao lado do arquivo
  `EpimedCultura.exe`.

# Epimed Cultura

Sistema para converter os relatórios do Epimed em Word e cruzá-los com a
planilha de Cultura.

## Como usar

1. Inicie o sistema de uma das seguintes formas:
   - Execute `python src\main.py` a partir da pasta do projeto; ou
   - Abra `dist\Cultura_Epimed.exe`, depois que o executável tiver sido gerado.
2. O sistema cria as pastas de trabalho que ainda não existem, sem apagar
   arquivos existentes.
3. Clique em **Conversões** e escolha **Word para Excel**,
   **Excel (xls) para Excel** ou **CSV para Excel**. O fluxo de **PDF para
   Excel** está em construção.
4. Selecione um ou mais arquivos do formato escolhido. O processamento começa
   automaticamente assim que você confirmar a seleção no botão **Abrir**. O
   seletor permite visualizar arquivos Word (`.docx`), Excel (`.xlsx`),
   Excel (xls), PDF e CSV. Word e CSV são salvos em uma subpasta `xlsx`; Excel
   (xls) é salvo na pasta dos arquivos selecionados. PDF ainda não pode ser
   processado.
5. Clique em **Gerar Base Final**:
   - Selecione um ou mais arquivos Excel (`.xlsx`) de Cultura.
   - Escolha o campo A da Cultura.
   - Selecione um ou mais arquivos Excel (`.xlsx`) do Epimed e escolha o campo
     B. Os seletores abrem inicialmente nas pastas conhecidas nesta sessão.
   - Escolha a pasta onde deseja salvar a base final. O seletor abre
     inicialmente na pasta `Saida`.
6. Os arquivos e abas são cruzados por mês; o campo A da Cultura e o campo B
   do Epimed podem ter nomes diferentes, mas cada campo deve existir nos
   arquivos e abas usados no cruzamento.
7. O resultado é salvo na pasta escolhida com o nome
   `Cultura_Epimed_DD_MM_YYYY.xlsx` (data atual), com uma aba por mês.

## Estrutura de pastas

```text
src\
  main.py                                Inicializador de desenvolvimento/PyInstaller
  cultura_epimed\
    ui\                                  Interface gráfica
    services\                            Conversão, pastas e cruzamento
    domain\                              Regras e estruturas de domínio
    assets\                              Recursos visuais do aplicativo
    epimed_culture_system.py             Fachada EpimedCultureSystem
```

As classes ficam agrupadas por responsabilidade dentro do pacote
`cultura_epimed`; não é necessário criar uma pasta para cada classe.

## Pastas de dados

```text
Culturas\             Planilha de Cultura no formato .xlsx
Dados_Epimed\         Relatórios do Epimed no formato .docx
Dados_Epimed\xlsx\    Arquivos convertidos pelo sistema
Saida\                Base final gerada pelo sistema
```

## Atenção

- A tela de abertura usa o logo oficial do
  [Projeto Impacto MR](https://www.bp.org.br/institucional/impacto-social/projetos-proadi-sus/trienio-2024-2026/impacto-mr).
- Para executar pelo módulo instalado, use `python -m pip install -e .` e
  depois `cultura-epimed`. Para executar diretamente o código-fonte, use
  `python src\main.py`.
- Para gerar o executável, instale os extras necessários com
  `python -m pip install -e ".[xls,build]"` e execute:

  ```powershell
  pyinstaller --onefile --windowed --name Cultura_Epimed --icon "src\cultura_epimed\assets\cultura-epimed.ico" --add-data "src\cultura_epimed\assets\logo-impacto-mr.png;assets" --add-data "src\cultura_epimed\assets\cultura-epimed.ico;assets" src\main.py
  ```
- O extra `xls` instala `xlrd`, necessário para conversão de arquivos XLS.
- Na conversão de XLS, o sistema detecta o formato real do arquivo e mantém
  todas as abas. Isso também aceita arquivos XLSX identificados incorretamente
  com a extensão `.xls`. Se a primeira linha produzir colunas `Unnamed`, o
  conversor usa a segunda linha como cabeçalho.
- Na leitura das planilhas de Cultura e Epimed, o sistema compara cabeçalhos
  na primeira e na segunda linha e usa a linha com mais nomes de coluna
  reconhecidos. Isso permite abas com cabeçalhos logo na primeira linha ou
  abaixo de um título.
- Se um arquivo Epimed parecer não ter cabeçalho, o sistema só infere os nomes
  quando pelo menos dois outros meses têm os mesmos cabeçalhos na mesma ordem
  e a mesma quantidade de colunas. Caso contrário, interrompe o fluxo e pede
  a correção ou reconversão do arquivo, sem arriscar associar campos errados.
- A conversão PDF para Excel ainda não está disponível.
- A coluna A escolhida precisa existir nos arquivos e abas de Cultura. A coluna
  B escolhida precisa existir nos arquivos e abas Epimed correspondentes.
- Cada mês/aba de Cultura precisa ter um arquivo ou uma aba Epimed correspondente
  entre os arquivos selecionados.
- Ao usar o executável, as pastas de trabalho serão criadas ao lado do arquivo
  `Cultura_Epimed.exe`.

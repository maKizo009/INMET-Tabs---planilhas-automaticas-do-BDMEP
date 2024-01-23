# **SHEETMASTERS** - Gere tabelas de dados históricos do INMET em poucos segundos!

Com o programa, você economiza um tempo precioso. A atividade repetitiva de copy + paste que durava horas agora é feita quase que instantaneamente
Ele foi desenvolvido usando Python, Pandas e a biblioteca padrão de interface gráfica Tkinter.

## Funções do programa:
- Cria tabelas prontas em 30 segundos
- Renomeia os arquivos finais para o nome da estação atualizada
- Preenche os dados automaticamente

## **Instruções**

A versão Windows do arquivo pode ser baixada [aqui](dist/auto_plan_inmet.exe).
Na página seguinte, clique no símbolo de download no canto centro-superior direito para baixar o arquivo .exe.
1- O programa, ao menos por enquanto, funciona apenas com arquivos csv baixados do site (https://bdmep.inmet.gov.br/).
  Arquivos de outros sites, ou até mesmo csv da tabela do mapa das estações ainda não são compatíveis e podem quebrar o código
2- Como medida de segurança, o Windows vai impedir o usuário de abrir o executável. Isso é o que se espera, no entanto, pode executar o programa normalmente. Como ele é de código aberto,
   caso o usuário queira, pode gerar uma nova compilação por conta própria, ou fiscalizar o código fonte aqui no Github.
3- Ao abrir o programa, uma "tela preta" irá aparecer. Isso é normal, é o terminal do Windows. Não a feche enquanto o programa não terminar de executar.
4- **Importante**: O arquivo .csv do BDMEP precisa ter seus dados de casas decimais separados por ponto ( . ). Caso use vírgula, ele irá falhar.
  O programa fará a conversão para ponto na criação da tabela, então não há problema.
5- Também vale salientar que você precisa **descompactar** a pasta dos arquivos csv antes de abrir. Caso contrário, o programa irá crashar.


Enfim, essas são as dicas e instruções principais para o bom uso do programa. Dediquei vários e vários meses para chegar a essa versão, e essa é a minha pequena contribuição para a meteorologia brasileira.
  


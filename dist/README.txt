------------------------------------------------------- Manual de Utilização do Script de Automatização de Cadastro -------------------------------------------------------

1 - O diretório em que se encontra o executavel deve conter os seguintes arquivos:
    1.1 - Cadastro.xlsx: Planilha de referência onde os dados referentes aos itens devem ser inseridos.
    1.2 - chromedriver.exe: Driver para acesso do google chrome através do package python. 
        1.2.1 - Nota importante: Esse arquivo pode ser incompatível com a sua versão do chrome, caso seja é necessário fazer o download da versão compatível em https://chromedriver.chromium.org/downloads 

2 - Para começar a fazer o cadastro, basta executar o arquivo "CadastroPortalCompras.exe" e fornecer os dados do usuário
3 - A depender do desempenho do computador em que esta sendo executado é possível utilizar a máquina normalmente enquanto o bot faz o trabalho, mas preferencialmente não utilize o computador durante o processo.
4 - Enquanto os itens vão sendo cadastrados,  o código do item no portal será anotado na planilha.
    4.1 - Nota importante: Manter a planilha fechada enquanto o processo é realizado.
5 - O bot realizará o cadastro completo do item, incluindo o envio para aprovação.

---------------------------------------------------------- Manual de Utilização do Script de Automatização de SC ----------------------------------------------------------

1 - O diretório em que se encontra o executavel deve conter os seguintes arquivos:
    1.1 - Cadastro.xlsx: Planilha de referência onde os dados referentes aos itens devem ser inseridos.
    1.2 - chromedriver.exe: Driver para acesso do google chrome através do package python. 
        1.2.1 - Nota importante: Esse arquivo pode ser incompatível com a sua versão do chrome, caso seja é necessário fazer o download da versão compatível em https://chromedriver.chromium.org/downloads 
2 - Para começar a fazer o cadastro, basta executar o arquivo "SolicitarCompra.exe" e fornecer os dados do usuário
3 - A depender do desempenho do computador em que esta sendo executado é possível utilizar a máquina normalmente enquanto o bot faz o trabalho, mas preferencialmente não utilize o computador durante o processo.
4 - O bot fará a abertura de uma nova SC, preenchendo os dados necessários e inserindo os itens na Planilha (com a quantidade de cada item)
    4.1 - Cabe ao usuário preencher a aba de "dados complementares" e anexar os arquivos necessários.


------------------------------------------------------------------ Fluxo de Utilização Completa dos BOTs ------------------------------------------------------------------

1 - O usuário deve inserir os itens que serão cadastrados na Planilha de cadastro, preenchendo corretamente as colunas e manter a coluna de "código do portal" vazia.
2 - Executar o "CadastroPortalCompras.exe" primeiro:
    2.1 - Assim que finalizado, verificar se os dados adicionados na planilha são coerentes.
    2.2 - Salvar a planilha.
3 - Executar o "SolicitarCompra.exe" em seguida:
    3.1 - Ao final, ainda será necessário acessar a SC pelo portal e inserir alguns dados complementares e anexar os arquivos referentes.

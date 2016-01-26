Este documento descreve a instala��o de dois programas para tratar os dados ADS-B (sinais de transponders dos avi�es comerciais) a partir do programa 'dump1090':

	- 'Dump1090sql�, armazena as informa��es em uma base SQLite3 e/ou arquivos de texto.
	- 'Dump1090report' , cria relat�rios em CSV (compat�vel com os arquivos Excel).

Os dois primeiros programas, 'dump1090' e 'dump1090sql' podem ser instalados como servi�os Windows e ser�o executados automaticamente ao reiniciar o PC. A grava��o dos dados nos arquivos textos e na base de dados ser� realizada 24h/24h. Um sistema de compress�o (formato .zip) e envio autom�tico de arquivos por e-mail nos finais de meses est� inclu�do nestes programas. Estes dois programas podem ser instalados em duas m�quinas remotas, basta que estejam na mesma rede (independentemente do tipo de O.S.).
O programa 'dump1090sql' funciona 24h/24h para alimentar as bases de dados e/ou os arquivos. O programa 'dump1090report� ser� utilizado nos finais de meses para gerar relat�rios a partir das informa��es dos v�os (dados dump1090sql), arquivos de alertas das esta��es de monitoramento de ru�dos DUO ou Cube (arquivos no formato .xls Excel) e do banco de dados 'basestation.sqb� que cont�m informa��es sobre aeronaves e n�meros de v�os.
Esses programas funcionam nos ambientes Windows, Linux ou OS X e utilizam scripts Python (vers�o > 2.7) e banco de dados SQLite3. O programa 'dump1090' corresponde a vers�o original n�o modificada, apenas as p�ginas HTML p�blicas do servidor embutido no programa foram modificadas para adicionar novos recursos (c�rculos de distancias, links para sites externos... ).

Na vers�o Windows apresentado aqui, o script Python foi compilado em um �nico arquivo execut�vel .exe para maior praticidade. Todos os arquivos devem estar na pasta 'c:\ dump1090sql' (modificar os arquivos de configura��o .ini para trocar de pasta).
Baixe e descompacte o arquivo 'dump1090-win� na mesma pasta (vers�o Windows do programa Linux DUMP1090). Pode-se tamb�m instalar o programa SQLiteStudio para verificar o conte�do das bases de dados, de prefer�ncia em uma subpasta (http://sqlitestudio.pl/).


'dump1090' foi escrito por Salvatore Sanfilippo antirez@gmail.com e esta sendo distribu�do sobre licen�a �BSD three clause license�.

SUMARIO

01. ESTRUTURA DA PASTA C:\DUMP1090SQL			4
02. INSTALA��O DOS DRIVERS ZADIG E UTILITARIOS 	5
03. INSTALA��O DO SERVI�O DUMP1090 				5
04. INSTALAR O SERVI�O DUMP1090SQL 				7
05. VERIFICA��O 								8
06. ARQUIVO �DUMP1090SQL.INI'					9
07. ESTRUTURA DAS PASTAS						10
08. CRIAR UM RELATORIO							11
09. FORMATO DOS ARQUIVOS .xls �ALARMA RU�DOS" 	16
10. DUMP1090SQL COM A OP��O "INTRACTIVE" 		16
11. AJUDA EM LINHA DE COMANDO 					17
12. SERVIDOR DUMP1090							18

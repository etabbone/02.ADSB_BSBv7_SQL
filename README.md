# 02.ADSB_BSBv7_SQL
**Python script to decode, save into database and create report from an DUMP1090 JSON stream (ADS-B Mode S decoder).**

Este documento descreve a instalação de dois programas para tratar os dados ADS-B (sinais de transponders dos aviões comerciais) a partir do programa 'dump1090':

	- 'Dump1090sql’, armazena as informações em uma base SQLite3 e/ou arquivos de texto.
	- 'Dump1090report' , cria relatórios em CSV (compatível com os arquivos Excel).

Os dois primeiros programas, 'dump1090' e 'dump1090sql' podem ser instalados como serviços Windows e serão executados automaticamente ao reiniciar o PC. A gravação dos dados nos arquivos textos e na base de dados será realizada 24h/24h. Um sistema de compressão (formato .zip) e envio automático de arquivos por e-mail nos finais de meses está incluído nestes programas. Estes dois programas podem ser instalados em duas máquinas remotas, basta que estejam na mesma rede (independentemente do tipo de O.S.).
O programa 'dump1090sql' funciona 24h/24h para alimentar as bases de dados e/ou os arquivos. O programa 'dump1090report’ será utilizado nos finais de meses para gerar relatórios a partir das informações dos vôos (dados dump1090sql), arquivos de alertas das estações de monitoramento de ruídos DUO ou Cube (arquivos no formato .xls Excel) e do banco de dados 'basestation.sqb’ que contém informações sobre aeronaves e números de vôos.
Esses programas funcionam nos ambientes Windows, Linux ou OS X e utilizam scripts Python (versão > 2.7) e banco de dados SQLite3. O programa 'dump1090' corresponde a versão original não modificada, apenas as páginas HTML públicas do servidor embutido no programa foram modificadas para adicionar novos recursos (círculos de distancias, links para sites externos... ).

Na versão Windows apresentado aqui, o script Python foi compilado em um único arquivo executável .exe para maior praticidade. Todos os arquivos devem estar na pasta 'c:\ dump1090sql' (modificar os arquivos de configuração .ini para trocar de pasta).
Baixe e descompacte o arquivo 'dump1090-win’ na mesma pasta (versão Windows do programa Linux DUMP1090). Pode-se também instalar o programa SQLiteStudio para verificar o conteúdo das bases de dados, de preferência em uma subpasta (http://sqlitestudio.pl/).


'[dump1090](https://github.com/MalcolmRobb/dump1090)' foi escrito por Salvatore Sanfilippo antirez@gmail.com e esta sendo distribuído sobre licença ‘BSD three clause license’.

![alt tag](https://github.com/etabbone/02.ADSB_BSBv7_SQL/blob/master/fluxo%20de%20dados.jpg)

**Sumário** ([dump1090sql](https://github.com/etabbone/02.ADSB_BSBv7_SQL/blob/master/dump1090sql.doc))

01. ESTRUTURA DA PASTA C:\DUMP1090SQL, 4
02. INSTALAÇÃO DOS DRIVERS ZADIG E UTILITARIOS, 5
03. INSTALAÇÃO DO SERVIÇO DUMP1090, 5
04. INSTALAR O SERVIÇO DUMP1090SQL, 7
05. VERIFICAÇÃO, 8
06. ARQUIVO 'DUMP1090SQL.INI', 9
07. ESTRUTURA DAS PASTAS, 10
08. CRIAR UM RELATORIO, 11
09. FORMATO DOS ARQUIVOS .xls “ALARMA RUÍDOS", 16
10. DUMP1090SQL COM A OPÇÃO "INTRACTIVE", 16
11. AJUDA EM LINHA DE COMANDO, 17
12. SERVIDOR DUMP1090, 18

*Ler "<b>[dump1090sql](https://github.com/etabbone/02.ADSB_BSBv7_SQL/blob/master/dump1090sql.doc)</b>" para mais informações* (em francês!!! mas é fácil de entender...).


**Command line**
<p>$ dump1090sql --help</p>

*Usage:*
<p>$ dump1090sql [option(s)]</p>

*Options:*
```
  -h, --help                 Show this help message and exit
      --ip "ip"              Set HTTP server IP (default: 127.0.0.1)
      --port "port"          Set HTTP server port (default: 8080)
      --delay "seconds"      Set delay before reconnecting on lost connexion
  -s, --sql                  Insert into SQLite database
      --sqlfile "database"   Set SQLite database filename
  -t, --txt                  Write to txt file
      --txtfile "file"       Set txt output filename
  -e, --exclude              Exclude inconsistent data
      --maxalt "altitude"    Set maxalt altitude flight
      --maxdist "distance"   Set maximum distance between flight and station
      --ifile "file"         Load Excel "file" into database (need --sqlfile)
      --report               Make report from database (need --sqlfile)
      --extdb "database"     Add flights infos (need --report or --ireport)
      --ireport "file"       Add flights infos to report "file" (need --extdb)
      --opencsv              Open csv report file (need --report)
      --zip "file"           Compress all files in one "zip file"
      --email                Send end of month email, (need --zip)
      --remove               Remove files after compression (need --zip)
      --lat "latitude"       Set latitude of local station (DD)
      --lon "longitude"      Set longitude of local station (DD)
      --alt "altitude"       Set altitude of local station (meters)
      --sta "name"           Set local station name
      --metric               Use metric units (meters, km/h, ...)
  -q, --quiet                Disable all output
  -i, --interactive          Show all informations
      --pause                Enable pauses during process
      --log                  Create a log file
  -v, --version              Show software version

Type CTRL+C to exit
```

## Projeto Chachá
Esse projeto foi criado para automatizar o processo de confecção de crachá, baseado em resultados obtidos via Google Forms.


### Requerimentos

Para poder fazer a comunicação com o Google Forms, será necessário a configuração de uma API, a ser utilizada como client_secret.json
Após o primeiro logon com sucesso, um token será gerado.

`SAMPLE_SPREADSHEET_ID = "insira aqui o ID da planilha"`

`SAMPLE_RANGE_NAME = "nome_da_aba!A:Z"`

Tamanho do crachá a ser gerado

`LARGURA = 100`

`ALTURA = 145`

Modelo do crachá sem os dados em png/jpg na pasta do codigo 

`modelocracha.png`

### Popular dados
Para pegar os dados da planilha, o código se adequa da seguinte maneira:

`variavel = row[posição da coluna no spreadsheet]`

Você pode ajustar de forma que desejar.

O sistema também tem um kerning padrão, ajustado de acordo com o tamanho da letra utilizada.



# Tela de Espera para Eventos

Aplicacao em Python para exibir uma tela de espera institucional em eventos, formacoes e momentos de organizacao do espaco.

Ela mostra relogio em tempo real, dados do evento, mensagem de boas-vindas, reproduz audio local e permite escolher fundos visuais, incluindo um fundo animado com a **Ilha do Pescador**.

## O que o projeto faz

- Exibe a tela em modo cheio com identidade visual personalizavel.
- Permite informar nome do evento, data, professores e mensagem.
- Preenche a data atual automaticamente no launcher, com possibilidade de edicao.
- Reproduz audio local com backend mais estavel.
- Mostra alertas visuais programados 5 minutos antes de horarios importantes.
- Oferece modo de teste para os avisos.
- Permite escolher fundo estatico ou fundo animado.

## Destaques atuais

- **Audio local mais estavel** com `pygame-ce`, com fallback para Windows Media Player e PowerShell quando necessario.
- **Avisos animados** com confetes e foguetes nos horarios programados.
- **Melhorias de legibilidade** nos textos da tela principal.
- **Fundo "Ilha do Pescador"** com:
  - sol se movimentando conforme o horario do dia
  - reflexo do sol acompanhando a agua
  - nuvens em deriva suave
  - pescador com animacoes calmas, rosto orientado pela acao e caminhada ate o coqueiro
  - pescarias engracadas com peixe, peixinho, bota, bigorna e pneu
  - item arremessado para a ilha e exibido no chao
  - gaivota que pode levar o peixe embora
  - tubarao lento no mar
  - caranguejo que sobe pela ilha
  - pescador fugindo para o meio da ilha quando o tubarao se aproxima
  - pescador balancando o coqueiro, recebendo e consumindo um coco
  - aviao distante observado com binoculo
  - cardume de peixinhos e golfinhos saltando ao fundo

## Horarios dos avisos

Os avisos sao disparados **5 minutos antes** destes horarios:

- `11:55` - aviso de almoco
- `16:55` - termino do expediente do primeiro professor
- `17:55` - termino do expediente do segundo professor

No uso normal, aparece apenas a animacao para chamar atencao.
No launcher existe tambem um **modo de teste** para visualizar o aviso imediatamente.

## Requisitos

- Windows 10 ou Windows 11
- Python 3.10 ou superior
- `tkinter` disponivel na instalacao do Python

## Dependencias

Instale com:

```bash
pip install -r requirements.txt
```

Pacotes usados:

- `Pillow` para tratamento de imagens
- `pygame-ce` para reproducao de audio local
- `pywin32` para integracao com recursos do Windows

## Como executar

```bash
python main.py
```

## Como usar

1. Abra o launcher.
2. Preencha ou ajuste os dados do evento.
3. Escolha o espaco, o fundo e o audio.
4. Se quiser, use o botao de teste de aviso.
5. Ajuste o intervalo do tubarao se desejar.
6. Preencha os agendamentos de banner aereo que desejar.
7. Clique em `Iniciar tela de espera`.
8. Para sair da tela cheia, pressione `Esc`.

## Mensagens Agendadas

O launcher permite cadastrar ate 8 mensagens em horarios definidos.

- Preencha o horario no formato `HH:MM`; ao digitar `1327`, o campo formata para `13:27`.
- Preencha a mensagem ao lado do horario.
- Linhas vazias, incompletas ou com horario invalido sao ignoradas.
- Cada mensagem e exibida uma vez por dia, no minuto configurado.
- Um aviao atravessa a faixa superior da tela puxando um banner de alto contraste.
- A passagem dura 30 segundos; se houver mais de uma mensagem no mesmo minuto, elas entram em fila.

## Fundos disponiveis

- `Padrao do espaco`
- `Fundo 1` a `Fundo 6` quando os arquivos existirem em `assets/`
- `Ilha do Pescador`

## Fundo Ilha do Pescador

O fundo animado da ilha e desenhado por codigo e nao depende de imagem extra.

Ele inclui:

- variacao de ceu, mar e sol ao longo do dia
- pescador no trapiche
- arremesso do item pescado para a areia da ilha
- cena do coqueiro: caminhada, balanco, coco caindo na mao e consumo
- cena de aviao distante com binoculo
- tubarao no mar com intervalo configuravel
- caranguejo na areia usando o mesmo intervalo configurado do tubarao
- gaivota que pode levar o peixe embora
- peixinhos e golfinhos saltando aleatoriamente no mar

## Configuracao do tubarao

No launcher existe um campo para definir o intervalo do tubarao em minutos.

- Se o valor for `10`, o tubarao reaparece aproximadamente nesse intervalo.
- Se o valor for `0`, o tubarao fica desativado.
- O caranguejo usa esse mesmo intervalo base, com pequena defasagem para nao coincidir exatamente.

No `config.json`, o valor correspondente e:

```json
{
  "intervalo_tubarao_minutos": 10
}
```

## Estrutura do projeto

```text
projeto/
|-- main.py
|-- config.json
|-- requirements.txt
|-- build.bat
|-- README.md
`-- assets/
    |-- Logo-colorido.png
    |-- Futuro_em_Movimento_TAW.mp3
    |-- fundo1.png
    |-- fundo2.png
    |-- fundo3.png
    |-- fundo4.png
    |-- fundo5.png
    |-- fundo6.png
    `-- LEIA-ME.md
```

## Build do executavel

Para gerar o executavel:

```bat
build.bat
```

O script:

- instala as dependencias de build
- usa uma unidade virtual temporaria para evitar erro de caminho grande no Windows
- gera a pasta do executavel dentro de `dist\`
- copia `assets` e `config.json` junto com o executavel

Depois disso, distribua a **pasta inteira** gerada em `dist\`.

No computador de destino, basta executar o `.exe` gerado nessa pasta.

## Configuracao

O arquivo `config.json` guarda os valores padrao usados pelo launcher e pela tela.

Exemplo:

```json
{
  "evento": "Nome do Evento",
  "data": "20/08/2026",
  "mensagem_boas_vindas": "Texto exibido na tela principal.",
  "professores": ["Prof. Nome", "Profa. Nome"],
  "espaco_padrao": "Robotica e Modelagem",
  "radio_padrao": "Musica local (MP3)",
  "volume_padrao": 60,
  "audio_local": "assets\\arquivo.mp3",
  "fundo_padrao": "Ilha do Pescador",
  "intervalo_tubarao_minutos": 10,
  "agendamentos_banner": [
    {
      "horario": "13:30",
      "mensagem": "A atividade comeca em breve."
    }
  ]
}
```

## Observacoes

- O projeto foi pensado para uso em ambiente Windows.
- Se o audio nao puder usar o backend principal, o sistema tenta alternativas compativeis.
- O fundo animado do pescador e desenhado por codigo, entao nao depende de imagem de fundo extra.
- O `README` pode evoluir junto com a cena animada, ja que esse fundo tem recebido refinamentos frequentes.

## Autor

**Julio Cesar Valera**
Professor - Rede Estadual de Sao Paulo
`juliovalera@professor.educacao.sp.gov.br`

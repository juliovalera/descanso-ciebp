# Tela de Espera para Eventos

> Uma tela de espera animada para eventos, formações e momentos de organização do espaço.

Aplicação para Windows desenvolvida em Python. Exibe dados do evento, relógio em tempo real, música e cenários animados contínuos — tudo configurável por uma interface simples.

## Visão geral

| Personalização | Animação | Uso prático |
| --- | --- | --- |
| Evento, data, equipe, mensagem e cores | Cenários vivos, relógio e avisos visuais | Tela cheia para projetores e TVs |
| Opacidade, intermitência e posição do relógio | Música local, avião com banner e alertas | Configuração salva entre os usos |

## Recursos principais

- **Dados do evento**: nome, data, equipe responsável, espaço e mensagem de boas-vindas.
- **Data automática**: o campo começa com a data atual e pode ser alterado livremente.
- **Textos personalizáveis**: cada texto possui cor própria, opacidade e opção de intermitência.
- **Relógio flexível**: tamanho grande no centro ou versão compacta no canto inferior direito.
- **Áudio de fundo**: reprodução de MP3 com alternativas para maior compatibilidade no Windows.
- **Avisos programados**: animação chamativa cinco minutos antes de `12:00`, `17:00` e `18:00`.
- **Mensagens aéreas**: até oito mensagens podem ser programadas; um avião leva o banner pelo alto da tela.
- **Execução simples**: pressione `Enter` no launcher para iniciar a tela de espera ou `Esc` para sair dela.

## Cenários disponíveis

### Ilha do Pescador

Uma cena tranquila, desenhada por código, que evolui continuamente sem exigir imagens adicionais.

- Sol, céu e reflexo na água acompanham o horário do dia.
- Nuvens, peixes e golfinhos se movimentam suavemente.
- O pescador alterna ações: pesca itens divertidos, observa um avião, balança o coqueiro e consome um coco.
- Uma gaivota pode buscar o peixe pescado; tubarão e caranguejo aparecem em intervalos configuráveis.

### Mundo de Blocos

Cena original em linguagem visual voxel, também criada por código.

- Sol quadrado, lua e estrelas conforme o período do dia.
- Construtor caminha pelo terreno e trabalha em obras completas, como cabana, oficina, torre, ponte e mirante.
- Animais, horta, riacho, partículas de construção e elementos do céu mantêm o cenário vivo.
- Uma criatura verde pode danificar uma obra durante o dia; o construtor prioriza o reparo antes de continuar.

### Voo no Rio

Uma animação de voo vertical em pixel art, sem combate e sem condição de derrota.

- Aeronave com piloto automático, deslocamentos laterais e verticais para evitar obstáculos.
- Rio, margens, barcos, helicópteros, ilhas e elementos de resgate em movimento contínuo.
- Medidor de **combustível** em estilo arcade e rota de reabastecimento acessível antes do fim da reserva.
- Ambiente leve e seguro: a navegação prossegue sem colisões ou explosões.

## Mensagens agendadas no avião

Cadastre até oito mensagens para horários específicos.

1. Informe o horário no formato `HH:MM`. Ao digitar `1327`, o campo formata automaticamente para `13:27`.
2. Escreva a mensagem ao lado do horário.
3. No momento programado, o avião atravessa a parte superior da tela puxando um banner de alto contraste.

Cada passagem leva aproximadamente 30 segundos. Linhas vazias, incompletas ou com horário inválido são ignoradas; mensagens do mesmo minuto entram em fila.

## Textos, visibilidade e acessibilidade

Os elementos abaixo podem ser configurados individualmente:

| Elemento | Cor | Opacidade | Intermitência | Ocultar |
| --- | :---: | :---: | :---: | :---: |
| Hora | Sim | Sim | Sim | Sim |
| Nome do evento | Sim | Sim | Sim | Sim |
| Legenda do espaço | Sim | Sim | Sim | Sim |
| Data | Sim | Sim | Sim | Sim |
| Professores | Sim | Sim | Sim | Sim |
| Mensagem | Sim | Sim | Sim | Sim |

As opções de opacidade são `100%`, `75%`, `50%` e `25%`. A intermitência oferece intervalos de segundos e minutos, além de **Desligado** e **Ocultar**.

## Avisos de horário

Os avisos são iniciados cinco minutos antes dos horários a seguir:

| Horário do aviso | Finalidade |
| --- | --- |
| `11:55` | Preparação para o almoço |
| `16:55` | Término do expediente do primeiro professor |
| `17:55` | Término do expediente do segundo professor |

O launcher também possui um modo de teste para conferir a animação imediatamente.

## Requisitos

- Windows 10 ou Windows 11
- Python 3.10 ou superior
- `tkinter` disponível na instalação do Python

## Instalação e execução

```bash
pip install -r requirements.txt
python main.py
```

Dependências utilizadas:

- `Pillow` — tratamento de imagens
- `pygame-ce` — reprodução de áudio local
- `pywin32` — integração com recursos do Windows

## Como usar

1. Abra o launcher com `python main.py`.
2. Preencha ou revise os dados do evento.
3. Escolha o espaço, cenário e áudio de fundo.
4. Ajuste cores, opacidade, intermitência e agendamentos, se desejar.
5. Clique em **Iniciar tela de espera** ou pressione `Enter`.
6. Na tela cheia, pressione `Esc` para sair.

## Gerar o executável

No Windows, execute:

```bat
build.bat
```

O script instala as dependências de compilação, evita o problema de caminhos longos do Windows e gera a distribuição dentro de `dist\\`.

> Para distribuir, copie a pasta inteira criada em `dist\\`. No computador de destino, basta abrir o arquivo `.exe` dentro dessa pasta.

## Configuração

O arquivo `config.json` armazena os valores utilizados pelo launcher. Exemplo resumido:

```json
{
  "evento": "Nome do Evento",
  "data": "20/08/2026",
  "mensagem_boas_vindas": "Texto exibido na tela principal.",
  "professores": ["Prof. Nome", "Profa. Nome"],
  "fundo_padrao": "Ilha do Pescador",
  "intervalo_tubarao_minutos": 10,
  "agendamentos_banner": [
    {
      "horario": "13:30",
      "mensagem": "A atividade começa em breve."
    }
  ]
}
```

Para desativar o tubarão da Ilha do Pescador, use `0` em `intervalo_tubarao_minutos`.

## Estrutura do projeto

```text
projeto/
├── main.py
├── config.json
├── requirements.txt
├── build.bat
├── README.md
└── assets/
    ├── Logo-colorido.png
    ├── Futuro_em_Movimento_TAW.mp3
    ├── fundo1.png … fundo6.png
    └── LEIA-ME.md
```

## Observações

- O projeto foi pensado para o ambiente Windows.
- Caso o áudio não possa usar o backend principal, o sistema tenta alternativas compatíveis.
- Os cenários animados são desenhados por código e não dependem de imagens extras.

## Autor

Julio Cesar Valera<br>
Professor da rede estadual de São Paulo<br>
`juliovalera@professor.educacao.sp.gov.br`

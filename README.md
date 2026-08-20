# CIEBP - Tela de Espera

Aplicação em Python para exibir uma tela de espera institucional do **CIEBP** em eventos, formações e momentos de organização do espaço.

Ela mostra relógio em tempo real, dados do evento, mensagem de boas-vindas, reproduz áudio local e permite escolher fundos visuais, incluindo um fundo animado com a **Ilha do Pescador**.

## O que o projeto faz

- Exibe a tela em modo cheio com identidade visual do CIEBP.
- Permite informar nome do evento, data, professores e mensagem.
- Preenche a data atual automaticamente no launcher, com possibilidade de edição.
- Reproduz áudio local com backend mais estável.
- Mostra alertas visuais programados 5 minutos antes de horários importantes.
- Oferece modo de teste para os avisos.
- Permite escolher fundo estático ou fundo animado.

## Destaques atuais

- **Áudio local mais estável** com `pygame-ce`, com fallback para Windows Media Player e PowerShell quando necessário.
- **Avisos animados** com confetes e foguetes nos horários programados.
- **Melhorias de legibilidade** nos textos da tela principal.
- **Fundo "Ilha do Pescador"** com:
  - sol se movimentando conforme o horário do dia
  - reflexo do sol acompanhando a água
  - pescador com animações calmas
  - pescarias engraçadas com peixe, peixinho, bota, bigorna e pneu
  - item arremessado para a ilha e exibido no chão
  - gaivota que pode levar o peixe embora
  - tubarão lento no mar
  - caranguejo que sobe pela ilha
  - pescador fugindo para o meio da ilha quando o tubarão se aproxima

## Horários dos avisos

Os avisos são disparados **5 minutos antes** destes horários:

- `11:55` - aviso de almoço
- `16:55` - término do expediente do primeiro professor
- `17:55` - término do expediente do segundo professor

No uso normal, aparece apenas a animação para chamar atenção.  
No launcher existe também um **modo de teste** para visualizar o aviso imediatamente.

## Requisitos

- Windows 10 ou Windows 11
- Python 3.10 ou superior
- `tkinter` disponível na instalação do Python

## Dependências

Instale com:

```bash
pip install -r requirements.txt
```

Pacotes usados:

- `Pillow` para tratamento de imagens
- `pygame-ce` para reprodução de áudio local
- `pywin32` para integração com recursos do Windows

## Como executar

```bash
python main.py
```

## Como usar

1. Abra o launcher.
2. Preencha ou ajuste os dados do evento.
3. Escolha o espaço, o fundo e o áudio.
4. Se quiser, use o botão de teste de aviso.
5. Ajuste o intervalo do tubarão se desejar.
6. Clique em `Iniciar tela de espera`.
7. Para sair da tela cheia, pressione `Esc`.

## Fundos disponíveis

- `Padrão do espaço`
- `Fundo 1` a `Fundo 6` quando os arquivos existirem em `assets/`
- `Ilha do Pescador`

## Fundo Ilha do Pescador

O fundo animado da ilha é desenhado por código e não depende de imagem extra.

Ele inclui:

- variação de céu, mar e sol ao longo do dia
- pescador no trapiche
- arremesso do item pescado para a areia da ilha
- tubarão no mar com intervalo configurável
- caranguejo na areia usando o mesmo intervalo configurado do tubarão
- gaivota que pode levar o peixe embora

## Configuração do tubarão

No launcher existe um campo para definir o intervalo do tubarão em minutos.

- Se o valor for `10`, o tubarão reaparece aproximadamente nesse intervalo.
- Se o valor for `0`, o tubarão fica desativado.
- O caranguejo usa esse mesmo intervalo base, com pequena defasagem para não coincidir exatamente.

No `config.json`, o valor correspondente é:

```json
{
  "intervalo_tubarao_minutos": 10
}
```

## Estrutura do projeto

```text
descanso-ciebp/
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

## Build do executável

Para gerar o executável:

```bat
build.bat
```

O script:

- instala as dependências de build
- usa uma unidade virtual temporária para evitar erro de caminho grande no Windows
- gera a pasta `dist\CIEBP_Descanso_de_Tela\`
- copia `assets` e `config.json` junto com o executável

Depois disso, distribua a **pasta inteira**:

```text
dist\CIEBP_Descanso_de_Tela\
```

No computador de destino, basta executar:

```text
CIEBP_Descanso_de_Tela.exe
```

## Configuração

O arquivo `config.json` guarda os valores padrão usados pelo launcher e pela tela.

Exemplo:

```json
{
  "evento": "Nome do Evento",
  "data": "20/08/2026",
  "mensagem_boas_vindas": "Texto exibido na tela principal.",
  "professores": ["Prof. Nome", "Profa. Nome"],
  "espaco_padrao": "Robótica e Modelagem",
  "radio_padrao": "Música local (MP3)",
  "volume_padrao": 60,
  "audio_local": "assets\\arquivo.mp3",
  "fundo_padrao": "Ilha do Pescador",
  "intervalo_tubarao_minutos": 10
}
```

## Observações

- O projeto foi pensado para uso em ambiente Windows.
- Se o áudio não puder usar o backend principal, o sistema tenta alternativas compatíveis.
- O fundo animado do pescador é desenhado por código, então não depende de imagem de fundo extra.
- O `README` pode evoluir junto com a cena animada, já que esse fundo tem recebido refinamentos frequentes.

## Autor

**Júlio César Valera**  
Professor - Rede Estadual de São Paulo  
CIEBP  
`juliovalera@professor.educacao.sp.gov.br`

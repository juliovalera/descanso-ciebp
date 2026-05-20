# CIEBP — Descanso de Tela Institucional

Tela de recepção profissional para eventos do **Centro de Inovação da Escola Básica Paulista (CIEBP)**.  
Exibe informações do evento, relógio em tempo real e toca música local enquanto aguarda o início das atividades.

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)
![Plataforma](https://img.shields.io/badge/Plataforma-Windows-informational?logo=windows)
![Licença](https://img.shields.io/badge/Licen%C3%A7a-MIT-green)

---

## Funcionalidades

- **Tela cheia institucional** com imagem de fundo por espaço do CIEBP
- **Relógio em tempo real** (hora e data atualizados a cada segundo)
- **Informações do evento**: nome, data e professores responsáveis
- **Mensagem de boas-vindas** configurável
- **Reprodução de MP3 local** via Windows Media Player (COM) ou PowerShell como fallback
- **Launcher de configuração** — janela inicial para ajustar o evento antes de exibir a tela
- **Fundo automático** em degradê institucional caso a imagem do espaço não seja encontrada
- **Portável**: funciona como script Python ou como executável `.exe` gerado via PyInstaller

---

## Espaços suportados

| Espaço                          | Cor de destaque |
|---------------------------------|-----------------|
| Hub de Inovação                 | `#00BCD4`       |
| Programação Descomplicada       | `#4CAF50`       |
| Cultura Maker                   | `#FF9800`       |
| Cultura Digital                 | `#9C27B0`       |
| Robótica e Modelagem            | `#F44336`       |
| Prototipação e Fabricação Digital | `#2196F3`     |

---

## Pré-requisitos

- **Windows 10 ou 11**
- **Python 3.10+** com tkinter (incluído na instalação padrão)

---

## Instalação

```bash
# 1. Clone o repositório
git clone https://github.com/juliovalera/descanso-ciebp.git
cd descanso-ciebp

# 2. Instale as dependências
pip install -r requirements.txt
```

### Dependências

| Pacote    | Finalidade                                      |
|-----------|-------------------------------------------------|
| `Pillow`  | Redimensionamento e sobreposição de imagens     |
| `pywin32` | Controle do Windows Media Player via COM (opcional) |

> **Nota:** sem `pywin32`, o programa usa o PowerShell embutido no Windows como fallback de áudio — nenhum programa externo é necessário.

---

## Como usar

### Rodando pelo Python

```bash
python main.py
```

1. A janela de configuração abrirá primeiro.
2. Preencha (ou confirme) os dados do evento: nome, data, professores e mensagem.
3. Selecione o **espaço** e o **fundo** desejado.
4. Clique em **Iniciar** — a tela cheia será exibida.
5. Pressione `Esc` para encerrar.

### Gerando o executável `.exe`

```bat
build.bat
```

O executável `CIEBP_Descanso_de_Tela.exe` será criado na pasta `dist/`.  
Copie também a pasta `assets/` e o arquivo `config.json` para o mesmo diretório do `.exe`.

---

## Estrutura do projeto

```
descanso-ciebp/
├── main.py            # Código principal (launcher + tela cheia)
├── config.json        # Configurações do evento (editável)
├── requirements.txt   # Dependências Python
├── build.bat          # Script para gerar o executável com PyInstaller
└── assets/
    ├── LEIA-ME.md                      # Instruções sobre as imagens
    ├── Logo-colorido.png               # Logo institucional do CIEBP
    ├── Futuro_em_Movimento_TAW.mp3     # Música de fundo padrão
    ├── fundo1.png … fundo6.png         # Imagens de fundo dos espaços
    └── (demais imagens por espaço)
```

---

## Configuração (`config.json`)

Todos os parâmetros do evento são editáveis diretamente no `config.json` ou pela janela de configuração do próprio programa:

```json
{
  "evento": "Nome do Evento",
  "data": "dd/mm/aaaa",
  "mensagem_boas_vindas": "Texto exibido na tela de recepção.",
  "professores": ["Prof. Nome", "Profa. Nome"],
  "espaco_padrao": "Cultura Digital",
  "radio_padrao": "Música local (MP3)",
  "volume_padrao": 60,
  "audio_local": "assets\\nome_do_arquivo.mp3"
}
```

---

## Imagens de fundo

Coloque imagens em `assets/` com as especificações abaixo e mapeie-as no `config.json`:

- **Resolução mínima:** 1920 × 1080 px (Full HD)
- **Formato:** PNG ou JPG
- **Orientação:** paisagem (horizontal)

Se uma imagem não for encontrada, o programa gera automaticamente um fundo em degradê escuro na cor de destaque do espaço.

---

## Licença

Este projeto está licenciado sob a [MIT License](LICENSE).

---

## Autor

**Júlio César Valera**  
Professor — Rede Estadual de São Paulo · CIEBP  
[juliovalera@professor.educacao.sp.gov.br](mailto:juliovalera@professor.educacao.sp.gov.br)

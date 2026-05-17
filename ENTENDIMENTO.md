# Peposoft — Simulador de Seleções

Aplicativo desktop em **Python + Tkinter** (`peposoft.py`) que simula partidas de futebol entre seleções nacionais, minuto a minuto, com narração ao vivo e visual temático.

## O que o app faz

- Catálogo de **20 seleções**, cada uma com:
  - `forca` (rating numérico, 76–95)
  - `flag` (emoji da bandeira)
  - `cor` (cor oficial em hex, usada na UI)
- Usuário escolhe **mandante × visitante** em dois `ttk.Combobox`.
- **Preview pré-jogo:** bandeiras, nomes coloridos e barra de probabilidade calculada pela razão das forças.
- **Simulação minuto a minuto** (1 → 90), usando `root.after` para criar ritmo:
  - Sorteia posse com base na força relativa.
  - ~18% de chance de virar finalização — chance de gol via `calcular_chance_gol`, escalada por 8.
  - Caso contrário, pode emitir um evento neutro de `EVENTOS_NEUTROS` (escanteio, falta, cartão, etc.).
- **Quando sai gol:** incrementa placar, faz *flash* visual no card com a cor do time e narra usando frase aleatória de `NARRACOES_GOL`.
- Marca intervalo aos 45' e fim aos 90'; ao fim, mostra vencedor ou empate.
- Botão **Resumo:** `messagebox` com placar final, posse de bola (%) e finalizações.

## Estrutura do código

Tudo em um único arquivo `peposoft.py`:

| Seção | Responsabilidade |
|---|---|
| `selecoes` | Dicionário com as 20 seleções e seus atributos |
| `NARRACOES_GOL`, `EVENTOS_NEUTROS` | Listas de frases para narração |
| Constantes de tema (`BG`, `ACCENT`, ...) | Paleta dark customizada |
| `calcular_chance_gol(f_atk, f_def)` | Probabilidade de gol por lance (clamp 0.005–0.08) |
| `probabilidades(f1, f2)` | % de força relativa entre os dois times |
| Classe `Simulador` | Toda a UI e lógica de simulação |

### Métodos da classe `Simulador`

- **Construção da UI** (privados): `_estilo`, `_cabecalho`, `_seletor`, `_barra_forca`, `_placar`, `_narracao`, `_botoes`
- **Preview:** `_atualizar_preview` — atualiza bandeiras, nomes e barra de probabilidade
- **Simulação:** `iniciar`, `_tick`, `_gol`, `_flash_placar`, `_restaurar_placar`, `_fim`
- **Auxiliares:** `_narrar` (escreve no widget `Text` com tags) e `mostrar_resumo`

## Aspectos técnicos notáveis

- **Tema dark customizado** com paleta no topo do arquivo; estilo `ttk` próprio (`Mod.TCombobox`) inclusive para a lista dropdown.
- **Barra de probabilidade** redesenhada no evento `<Configure>` da janela — acompanha redimensionamento.
- **Narração** em `tk.Text` com tags coloridas (`gol` em verde, `apito` em laranja).
- **Flash do placar** ao sair gol: pinta vários widgets com a cor do time por 220ms e restaura.
- Sem dependências externas — apenas stdlib (`random`, `tkinter`). Roda direto com `python peposoft.py`.

## Possíveis melhorias

- Pênaltis / prorrogação em caso de empate.
- Histórico de partidas simuladas.
- Modo torneio (grupos + mata-mata).
- Mais seleções ou edição manual de forças.
- Ajustes de balanceamento na fórmula de gol.

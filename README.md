# Clara Invest

Ferramenta em Excel que simula investimentos em fundos imobiliários (FIIs): do aporte de cada mês até o patrimônio acumulado e os dividendos mensais. Também projeta cenários de 2 a 30 anos e divide o aporte entre seis tipos de FII, conforme o perfil de investidor escolhido numa lista.

Projeto prático do desafio da DIO, com o Expert Felipe Aguiar (Felipão).

> Todos os valores da planilha (salário, aporte, taxas) são **exemplos**, não são dados reais.

## Prints

Mesma simulação (R$ 350/mês, 5 anos, 1,079% ao mês) em dois perfis:

| Perfil Conservador | Perfil Agressivo |
|---|---|
| ![Perfil conservador](perfil-conservador.png) | ![Perfil agressivo](perfil-agressivo.png) |

## Perguntas que a ferramenta responde

Todas na aba **APP**. As células brancas (texto azul) são entradas; as cinzas são calculadas.

| Pergunta | Célula |
|---|---|
| Quanto investir por mês? | `D17` (entrada) |
| Por quantos anos? | `D18` (entrada) |
| Qual a taxa de rendimento mensal? | `D19` (entrada) |
| Quanto de patrimônio vai acumular? | `D20` (calculado) |
| Quanto vai receber de dividendos por mês? | `D21` (calculado) |

Outras respostas na mesma aba:

- **Cenários de 2, 5, 10, 20 e 30 anos:** `C24:D28` (patrimônio e dividendos de cada prazo).
- **Simulação usando 30% do salário no lugar do aporte informado:** `D31:D34`.
- **Divisão do aporte por perfil:** perfil escolhido em `D37`, percentuais em `C41:C46`, valores em R$ em `D41:D46` e gráfico de pizza logo abaixo.

## Como o VF e o PROCV entram nos cálculos

### VF (`FV` no Excel em inglês)

O `VF` calcula o valor futuro de aportes mensais iguais, com juros compostos:

```
=VF(taxa_mensal; qtd_anos*12; aporte*-1)
```

- `taxa_mensal`: taxa de rendimento ao mês.
- `qtd_anos*12`: número de meses (o `VF` trabalha com períodos, então converto anos em meses).
- `aporte*-1`: o aporte entra negativo porque o Excel trata o dinheiro que sai do bolso como saída. Multiplicar por -1 faz o resultado vir positivo.

O resultado é o **patrimônio** (`D20`). Os **dividendos mensais** (`D21`) saem dele: `=patrimonio*rendimento_carteira`.

Nos cenários, repito o mesmo `VF` trocando só o prazo (2, 5, 10, 20 e 30 anos), com referências absolutas (`$D$19`, `$D$17`) para a fórmula não se perder ao arrastar:

```
=VF($D$19;$A24*12;$D$17*-1)
```

### PROCV (`VLOOKUP` no Excel em inglês)

A aba **Perfis** guarda uma tabela com perfil, tipo de FII e percentual. Cada linha tem uma **chave composta**, que junta perfil e tipo (por exemplo, `Moderado-TIJOLO`).

Na aba APP, o `PROCV` monta a mesma chave com o perfil escolhido na lista e o tipo de FII da linha, e traz o percentual:

```
=PROCV($D$37&"-"&B41;tabela_perfis;4;FALSO)
```

O valor em reais de cada tipo é o percentual vezes o aporte: `=C41*$D$38`.

A escolha do perfil usa **validação de dados** (lista suspensa), que aceita só Conservador, Moderado ou Agressivo.

## Intervalos nomeados

| Nome | Referência | O que é |
|---|---|---|
| `salario` | `APP!D12` | Salário (exemplo) |
| `rendimento_carteira` | `APP!D13` | Rendimento mensal da carteira em dividendos |
| `sugestao_investimento` | `APP!D14` | 30% do salário |
| `aporte` | `APP!D17` | Quanto investir por mês |
| `qtd_anos` | `APP!D18` | Por quantos anos |
| `taxa_mensal` | `APP!D19` | Taxa de rendimento mensal |
| `patrimonio` | `APP!D20` | Patrimônio acumulado |
| `dividendos` | `APP!D21` | Dividendos mensais |
| `perfil` | `APP!D37` | Perfil de investidor escolhido |
| `lista_perfis` | `Perfis!F3:F5` | Lista usada na validação de dados |
| `tabela_perfis` | `Perfis!A3:D20` | Tabela usada pelo `PROCV` |

## Percentuais de cada perfil

Os percentuais são **ilustrativos**, definidos por mim para fins didáticos. **Não são recomendação de investimento.** Cada perfil soma 100% (a aba Perfis confere isso e mostra "OK" ou "REVISAR").

| Tipo de FII | Conservador | Moderado | Agressivo |
|---|---|---|---|
| Papel | 40% | 30% | 20% |
| Tijolo | 35% | 35% | 25% |
| Híbridos | 10% | 10% | 10% |
| FOFs | 10% | 10% | 5% |
| Desenvolvimento | 0% | 10% | 25% |
| Hotelarias | 5% | 5% | 15% |
| **Total** | **100%** | **100%** | **100%** |

A lógica que usei: o perfil conservador concentra mais em papel e tijolo e não tem desenvolvimento; o agressivo aceita mais risco, com mais desenvolvimento e hotelarias.

## O que mudei em relação à ferramenta do Expert

- **Identidade própria:** logo e nome "Clara Invest", com paleta verde e dourada.
- **Percentuais próprios** em cada perfil, diferentes dos do Expert.
- **Simulação com 30% do salário** (`D31:D34`), que mostra também a diferença de patrimônio em relação ao aporte informado.
- **Conferência de 100%** por perfil na aba Perfis, além da linha de total na tabela da aba APP.
- **Mais intervalos nomeados:** `dividendos`, `perfil`, `lista_perfis` e `tabela_perfis`.
- **Lista de perfis ligada à aba Perfis** por intervalo nomeado, em vez de texto digitado na validação. Se eu incluir um perfil novo na tabela, a lista acompanha.
- **Rótulos dos cenários por fórmula** (`="Quanto em "&A24&" anos ?"`), que acompanham o número de anos.
- **Legenda** na planilha explicando o que é entrada e o que é calculado.
- **Valores de exemplo próprios:** aporte de R$ 350, salário de R$ 3.000 e taxa de 1,079% ao mês.

## Arquivos

- `Clara_Invest.xlsx`: a planilha (abas **APP** e **Perfis**).
- `perfil-conservador.png` e `perfil-agressivo.png`: prints da mesma simulação em dois perfis.

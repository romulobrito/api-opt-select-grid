# Análise Detalhada das Soluções de Otimização de Produção

## 1. Contextualização do Problema

### 1.1 Objetivo
Otimizar a produção de calças minimizando custos totais, considerando:
- Custos de tecido
- Custos operacionais (corte, setup, camadas)
- Desperdício de material
- Eficiência dos layouts

### 1.2 Parâmetros de Custo
- Preço do tecido por metro linear: R$18,90
- Custo de corte por metro: R$0,45
- Custo de setup por metro: R$0,30
- Custo por camada: R$1,25

### 1.3 Demanda
- P: 58 peças
- M: 79 peças
- G: 57 peças
- GG: 64 peças

## 2. Detalhamento dos Layouts

### 2.1 Características dos Layouts Disponíveis

| Layout | Utilização | Comprimento | Perímetro | Área Total | Área Útil | Desperdício | Produção por Layout |
|--------|------------|-------------|-----------|------------|-----------|-------------|---------------------|
| 1 | 82% | 7.772m | 91.88m | 11.6584m² | 9.5599m² | 2.0985m² | P:2, M:2, G:2, GG:1 |
| 2 | 85% | 7.001m | 82.11m | 10.5008m² | 8.9257m² | 1.5751m² | P:1, M:2, G:1, GG:2 |
| 3 | 94% | 7.921m | 106.05m | 11.8821m² | 11.1692m² | 0.7129m² | P:2, M:3, G:1, GG:2 |
| 4 | 91% | 7.233m | 93.35m | 10.8500m² | 9.8735m² | 0.9765m² | P:2, M:1, G:3, GG:1 |

## 3. Análise das Soluções

### 3.1 Solução A [10, 20, 5, 4]

#### Distribuição de Camadas e Produção
1. Layout 1 (82%): 10 camadas
   - Produção: P:20, M:20, G:20, GG:10
   - Tecido: 7.772m × 10 × R$18,90 = R$1.468,91
   - Corte: 91.88m × 10 × R$0,45 = R$413,46
   - Setup: (7.772m × R$0,30 × 10)/1000 + (R$1,25 × 10) = R$35,31

2. Layout 2 (85%): 20 camadas
   - Produção: P:20, M:40, G:20, GG:40
   - Tecido: 7.001m × 20 × R\$18,90 = R\$2.646,38
   - Corte: 82.11m × 20 × R\$0,45 = R\$738,99
   - Setup: (7.001m × R\$0,30 × 20)/1000 + (R\$1,25 × 20) = R\$67,00

3. Layout 3 (94%): 5 camadas
   - Produção: P:10, M:15, G:5, GG:10
   - Tecido: 7.921m × 5 × R\$18,90 = R\$748,53
   - Corte: 106.05m × 5 × R\$0,45 = R\$238,61
   - Setup: (7.921m × R\$0,30 × 5)/1000 + (R\$1,25 × 5) = R\$17,88

4. Layout 4 (91%): 4 camadas
   - Produção: P:8, M:4, G:12, GG:4
   - Tecido: 7.233m × 4 × R\$18,90 = R\$546,81
   - Corte: 93.35m × 4 × R\$0,45 = R\$168,03
   - Setup: (7.233m × R\$0,30 × 4)/1000 + (R\$1,25 × 4) = R\$13,67

#### Custos Totais Solução A
- Tecido: R$5.410,63
- Corte: R$1.559,09
- Setup: R$133,86
- Camadas: R$48,75
- Total: R$7.152,33

### 3.2 Solução B [12, 9, 16, 1] (Algoritmo)

#### Distribuição de Camadas e Produção
1. Layout 3 (94%): 12 camadas
   - Produção: P:24, M:36, G:12, GG:24
   - Tecido: 7.921m × 12 × R\$18,90 = R\$1.796,48
   - Corte: 106.05m × 12 × R\$0,45 = R\$57,27
   - Setup: (7.921m × R\$0,30 × 12)/1000 + (R\$1,25 × 12) = R\$43,52

2. Layout 4 (91%): 9 camadas
   - Produção: P:18, M:9, G:27, GG:9
   - Tecido: 7.233m × 9 × R\$18,90 = R\$1.230,33
   - Corte: 93.35m × 9 × R\$0,45 = R\$37,81
   - Setup: (7.233m × R\$0,30 × 9)/1000 + (R\$1,25 × 9) = R\$30,78

3. Layout 2 (85%): 16 camadas
   - Produção: P:16, M:32, G:16, GG:32
   - Tecido: 7.001m × 16 × R\$18,90 = R\$2.117,10
   - Corte: 82.11m × 16 × R\$0,45 = R\$59,12
   - Setup: (7.001m × R\$0,30 × 16)/1000 + (R\$1,25 × 16) = R\$53,60

4. Layout 1 (82%): 1 camada
   - Produção: P:2, M:2, G:2, GG:1
   - Tecido: 7.772m × 1 × R\$18,90 = R\$146,89
   - Corte: 91.88m × 1 × R\$0,45 = R\$4,13
   - Setup: (7.772m × R\$0,30 × 1)/1000 + (R\$1,25 × 1) = R\$3,58

#### Custos Totais Solução B
- Tecido: R$5.290,81
- Corte: R$158,33
- Setup: R$131,48
- Camadas: R$47,50
- Desperdício: R$562,52
- Total: R$6.190,63

## 4. Análise Comparativa

### 4.1 Métricas de Eficiência

| Métrica | Solução A | Solução B | Diferença |
|---------|-----------|-----------|------------|
| Metros de Tecido | 286,19m | 279,94m | -6,25m |
| Área Desperdiçada | 0,5897m² | 0,4464m² | -0,1433m² |
| Total de Camadas | 39 | 38 | -1 |
| Custo Total | R$7.152,33 | R$6.190,63 | -R$961,70 |

### 4.2 Identificação do Erro na Solução A

O erro principal na solução anterior estava na forma como os custos foram calculados:

1. **Erro no Cálculo do Tecido**:
   - Não considerava corretamente o número de camadas
   - Não aplicava a conversão correta de unidades (mm para m)
   - Resultado: R\$1.133,20 (incorreto) vs R\$5.410,63 (correto)

2. **Erro no Cálculo do Corte**:
   - Não multiplicava pelo número de camadas adequadamente
   - Resultado: R\$168,03 (incorreto) vs R\$1.559,09 (correto)

3. **Erro no Setup**:
   - Conversão incorreta de unidades
   - Resultado: R\$85,88 (incorreto) vs R\$133,86 (correto)

## 5. Conclusões

### 5.1 Vantagens da Solução B (Algoritmo)
1. Melhor utilização dos layouts mais eficientes
2. Menor consumo total de tecido
3. Menor área desperdiçada
4. Custos operacionais otimizados
5. Distribuição mais eficiente das camadas

### 5.2 Recomendações
1. Manter os parâmetros atuais do algoritmo
2. Considerar aumentar penalidades por:
   - Troca de layout
   - Desperdício de material
3. Avaliar aumento do número mínimo de camadas por layout
4. Implementar validações adicionais nos cálculos de custo

### 5.3 Impacto Financeiro
- Economia total: R$961,70 (13,4%)
- Redução no consumo de tecido: 6,25m (2,2%)
- Redução na área desperdiçada: 0,1433m² (24,3%)

## 6. Próximos Passos
1. Documentar metodologia de cálculo
2. Implementar verificações automáticas
3. Criar testes de validação
4. Monitorar resultados em produção

# Análise Comparativa das Três Soluções

## 7. Comparação das Soluções

### 7.1 Resumo das Soluções

| Aspecto | Solução A (Incorreta) | Solução A (Corrigida) | Solução B (Algoritmo) |
|---------|----------------------|---------------------|---------------------|
| Distribuição | [10, 20, 5, 4] | [10, 20, 5, 4] | [12, 9, 16, 1] |
| Custo Tecido | R$1.133,20 | R$5.410,63 | R$5.290,81 |
| Custo Corte | R$168,03 | R$1.559,09 | R$158,33 |
| Custo Setup | R$85,88 | R$133,86 | R$131,48 |
| Custo Camadas | R$48,75 | R$48,75 | R$47,50 |
| Custo Desperdício | - | R$589,70 | R$562,52 |
| **Custo Total** | **R$1.435,86** | **R$7.742,03** | **R$6.190,63** |

### 7.2 Análise por Layout

#### Layout 1 (82% utilização)
| Aspecto | Solução A (10 camadas) | Solução B (1 camada) |
|---------|----------------------|---------------------|
| Tecido | R$1.468,91 | R$146,89 |
| Corte | R$413,46 | R$4,13 |
| Setup | R$35,31 | R$3,58 |
| Desperdício | R$209,85 | R$20,99 |

#### Layout 2 (85% utilização)
| Aspecto | Solução A (20 camadas) | Solução B (16 camadas) |
|---------|----------------------|---------------------|
| Tecido | R$2.646,38 | R$2.117,10 |
| Corte | R$738,99 | R$59,12 |
| Setup | R$67,00 | R$53,60 |
| Desperdício | R$157,51 | R$126,01 |

#### Layout 3 (94% utilização)
| Aspecto | Solução A (5 camadas) | Solução B (12 camadas) |
|---------|----------------------|---------------------|
| Tecido | R$748,53 | R$1.796,48 |
| Corte | R$238,61 | R$57,27 |
| Setup | R$17,88 | R$43,52 |
| Desperdício | R$71,29 | R$171,10 |

#### Layout 4 (91% utilização)
| Aspecto | Solução A (4 camadas) | Solução B (9 camadas) |
|---------|----------------------|---------------------|
| Tecido | R$546,81 | R$1.230,33 |
| Corte | R$168,03 | R$37,81 |
| Setup | R$13,67 | R$30,78 |
| Desperdício | R$97,65 | R$219,71 |

### 7.3 Métricas de Eficiência

| Métrica | Solução A Corrigida | Solução B | Diferença |
|---------|-------------------|-----------|------------|
| Metros de Tecido | 286,19m | 279,94m | -6,25m (2,2%) |
| Área Desperdiçada | 0,5897m² | 0,4464m² | -0,1433m² (24,3%) |
| Total de Camadas | 39 | 38 | -1 (2,6%) |
| Custo Total | R$7.742,03 | R$6.190,63 | -R$1.551,40 (20,0%) |

### 7.4 Principais Diferenças

1. **Distribuição de Camadas**:
   - Solução A: Concentra mais camadas em layouts menos eficientes
   - Solução B: Prioriza layouts mais eficientes (94% e 91%)

2. **Custos Operacionais**:
   - Solução A: Maiores custos de corte devido à distribuição
   - Solução B: Otimiza operação
  
  # Análise de Eficiência e Impacto Financeiro

## 1. Redução de Custos Operacionais

### 1.1 Custos de Corte
- Redução significativa de 89,8% nos custos de corte
- Solução A: R$1.559,09
- Solução B: R$158,33
- Economia: R$1.400,76

### 1.2 Eficiência Material
| Aspecto | Solução A | Solução B | Melhoria |
|---------|-----------|-----------|-----------|
| Consumo de Tecido | 286,19m | 279,94m | 2,2% |
| Área Desperdiçada | 0,5897m² | 0,4464m² | 24,3% |

## 2. Impacto Financeiro Total

### 2.1 Economia por Categoria
| Categoria | Economia (R$) | Redução (%) |
|-----------|---------------|-------------|
| Tecido | 119,82 | 2,2% |
| Corte | 1.400,76 | 89,8% |
| Setup | 2,38 | 1,8% |
| Camadas | 1,25 | 2,6% |
| **Total** | **1.551,40** | **20,0%** |

## 3. Conclusões

### 3.1 Eficiência Operacional
- Redução de 20% nos custos totais
- Melhor utilização dos layouts mais eficientes
- Otimização da distribuição de camadas

### 3.2 Otimização de Recursos
- Redução de 6,25m no consumo de tecido
- Diminuição de 0,1433m² na área desperdiçada
- Distribuição mais eficiente das camadas entre layouts

### 3.3 Validação do Algoritmo
- Confirmação da eficácia do modelo de otimização
- Importância dos parâmetros de penalidade
- Justificativa para uso do modelo matemático

## 4. Recomendações

### 4.1 Ações Imediatas
1. Manter configuração atual do algoritmo
2. Implementar verificações sistemáticas de cálculo
3. Documentar detalhadamente a metodologia de custos

### 4.2 Monitoramento Contínuo
- Acompanhamento dos resultados em produção
- Validação periódica dos parâmetros
- Análise comparativa de desempenho

### 4.3 Melhorias Futuras
- Refinamento dos parâmetros de penalidade
- Desenvolvimento de indicadores de eficiência
- Automatização das verificações de qualidade

# Metodologia de Custos e Formulação Matemática para Otimização

## 1. Definições e Unidades Base

### 1.1 Unidades de Medida
| Dimensão   | Entrada | Conversão       | Uso                     |
|------------|---------|-----------------|-------------------------|
| Comprimento| mm      | ÷1000 → m       | Cálculo de tecido       |
| Perímetro  | cm      | ÷100 → m        | Cálculo de corte        |
| Área       | cm²     | ÷10\,000 → m²   | Cálculo de desperdício  |
| Largura    | mm      | ÷1000 → m       | Dimensionamento         |

### 1.2 Parâmetros de Custo
| Parâmetro                | Valor  | Unidade    | Aplicação            |
|--------------------------|--------|------------|----------------------|
| Preço por metro linear   | 18,90  | R\$/m      | Custo do tecido      |
| Custo por metro de corte | 0,45   | R\$/m      | Custo de corte       |
| Custo por metro de setup | 0,30   | R\$/m      | Custo de setup       |
| Custo por camada         | 1,25   | R\$/camada | Custo operacional    |
| Custo do desperdício por m² | 1.000 | R\$/m²  | Custo do desperdício |
| Penalidade por troca de layout | 2.500 | R\$ | Penalidade por troca de layout |
| Penalidade por superprodução   | 10.000 | R\$ | Penalidade por superprodução |

## 2. Formulação Matemática Detalhada

### 2.1 Definição das Variáveis de Decisão
- \( Y_l \): Variável binária que indica o uso do layout \( l \) (1 se usado, 0 caso contrário).
- \( S \): Quantidade de superprodução.

### 2.2 Funções de Custo

#### 2.2.1 Custos de Material

**Custo do Tecido**
\[
CT(l, c) = \frac{L_l \times P_m \times c}{1000}
\]
Onde:
- \( CT(l, c) \): Custo do tecido para o layout \( l \) com \( c \) camadas.
- \( L_l \): Comprimento do layout \( l \) em mm.
- \( P_m \): Preço por metro linear (R\$18,90).
- \( c \): Número de camadas.
- \( 1000 \): Fator de conversão de mm para m.

**Custo de Corte**
\[
CC(l, c) = \frac{P_l \times C_c \times c}{100}
\]
Onde:
- \( CC(l, c) \): Custo de corte para o layout \( l \) com \( c \) camadas.
- \( P_l \): Perímetro do layout \( l \) em cm.
- \( C_c \): Custo por metro de corte (R\$0,45).
- \( c \): Número de camadas.
- \( 100 \): Fator de conversão de cm para m.

#### 2.2.2 Custos Operacionais

**Custo de Setup**
\[
CS(l, c) = \frac{L_l \times C_s \times c}{1000} + (C_o \times c)
\]
Onde:
- \( CS(l, c) \): Custo de setup para o layout \( l \) com \( c \) camadas.
- \( L_l \): Comprimento do layout \( l \) em mm.
- \( C_s \): Custo por metro de setup (R\$0,30).
- \( c \): Número de camadas.
- \( C_o \): Custo operacional por camada (R\$1,25).

**Custo de Desperdício**
\[
CD(l, c) = \frac{A_d \times C_d \times c}{10\,000}
\]
Onde:
- \( CD(l, c) \): Custo do desperdício para o layout \( l \) com \( c \) camadas.
- \( A_d \): Área desperdiçada em cm².
- \( C_d \): Custo do desperdício por m² (R\$1.000).
- \( c \): Número de camadas.
- \( 10\,000 \): Fator de conversão de cm² para m².

### 2.3 Função Objetivo

A função objetivo visa minimizar os custos totais associados à produção, incluindo custos de material, custos operacionais e penalidades.

\[
\text{Min } Z = \sum_{l \in L} \sum_{c \in C} \left[ CT(l, c) + CC(l, c) + CS(l, c) + CD(l, c) \right] + \sum_{l \in L} P_l \times Y_l + P_s \times S
\]

Onde:
- \( L \): Conjunto de layouts.
- \( C \): Conjunto de camadas possíveis.
- \( P_l \): Penalidade por troca de layout (R\$2.500).
- \( Y_l \): Variável binária de uso do layout \( l \).
- \( P_s \): Penalidade por superprodução (R\$10.000).
- \( S \): Quantidade de superprodução.

### 2.4 Restrições

#### 2.4.1 Restrição de Utilização da Área
\[
\frac{\text{Área Útil}}{\text{Área Total}} \geq 0,80
\]
Onde:
- A utilização deve ser pelo menos 80%.

#### 2.4.2 Restrição de Superprodução
\[
\frac{\text{Produção} - \text{Demanda}}{\text{Demanda}} \leq 0,05
\]
Onde:
- A superprodução não deve exceder 5% da demanda.

#### 2.4.3 Restrição de Número de Camadas
\[
c \leq 50 \quad \forall c \in C
\]
Onde:
- O número de camadas por layout não pode exceder 50.

#### 2.4.4 Limites Operacionais
\[
\begin{cases}
L_l \leq 10\,000 \quad \forall l \in L \\
c \leq 30 \quad \forall c \in C \\
\text{Largura do Tecido} \leq 1\,500 \text{ mm}
\end{cases}
\]

## 3. Exemplo Numérico Detalhado

### 3.1 Cálculo para Layout 3 (12 camadas)

#### Dados do Layout
| Medida             | Valor Original | Unidade | Valor Convertido |
|--------------------|----------------|---------|------------------|
| Comprimento        | 7.921          | mm      | 7,921 m          |
| Perímetro          | 10.605         | cm      | 106,05 m         |
| Área Total         | 118.821        | cm²     | 11,8821 m²        |
| Área Desperdiçada  | 7.129          | cm²     | 0,7129 m²         |

#### Cálculos

1. **Custo do Tecido**
   \[
   CT = \frac{7.921 \times 18,90 \times 12}{1000} = R\$1.796,48
   \]

2. **Custo de Corte**
   \[
   CC = \frac{10.605 \times 0,45 \times 12}{100} = R\$57,27
   \]

3. **Custo de Setup**
   \[
   CS = \frac{7.921 \times 0,30 \times 12}{1000} + (1,25 \times 12) = R\$43,52
   \]

4. **Custo de Desperdício**
   \[
   CD = \frac{7.129 \times 1.000 \times 12}{10\,000} = R\$171,10
   \]

5. **Custo Total para o Layout 3**
   \[
   Z_3 = CT + CC + CS + CD = 1.796,48 + 57,27 + 43,52 + 171,10 = R\$2.068,37
   \]

## 4. Validações e Verificações

### 4.1 Checagem de Consistência
| Verificação    | Fórmula                                             | Critério  |
|----------------|-----------------------------------------------------|-----------|
| Utilização     | \(\frac{\text{Área Útil}}{\text{Área Total}}\)      | ≥ 80%     |
| Superprodução  | \(\frac{\text{Produção} - \text{Demanda}}{\text{Demanda}}\) | ≤ 5%      |
| Camadas        | \(\text{Número de Camadas}\)                        | ≤ 50      |

### 4.2 Limites Operacionais
| Parâmetro            | Limite | Unidade |
|----------------------|--------|---------|
| Comprimento Máximo   | 10.000 | mm      |
| Camadas por Layout   | 30     | unidades|
| Largura do Tecido    | 1.500  | mm      |

## 5. Análise de Sensibilidade

### 5.1 Impacto dos Parâmetros
| Parâmetro            | Variação | Impacto no Custo         |
|----------------------|----------|--------------------------|
| Preço do Tecido      | ±10%     | ±R\$179,65               |
| Custo de Corte       | ±10%     | ±R\$5,73                 |
| Penalidade Layout    | ±500     | ±R\$500 por troca        |

### 5.2 Pontos Críticos
- **Conversão de Unidades**: Erros na conversão podem impactar significativamente os custos.
- **Acumulação de Arredondamentos**: Pequenos erros podem se acumular, afetando a precisão dos resultados.
- **Precisão dos Cálculos**: Necessidade de manter alta precisão para evitar discrepâncias.
- **Validação Cruzada**: Verificar os cálculos por múltiplos métodos para garantir a exatidão.

## 6. Recomendações de Implementação

### 6.1 Boas Práticas
1. **Utilizar Tipos Decimais para Cálculos Monetários**: Evitar erros de arredondamento.
2. **Manter Precisão nas Conversões**: Usar fator de conversão exatos.
3. **Documentar Fatores de Conversão**: Facilitar auditorias e revisões.
4. **Validar Resultados Intermediários**: Garantir a precisão em cada etapa do cálculo.

### 6.2 Controles
1. **Verificação de Limites**: Assegurar que todas as restrições sejam atendidas.
2. **Validação de Resultados**: Comparar resultados com benchmarks ou dados históricos.
3. **Testes de Consistência**: Aplicar testes para verificar a lógica dos cálculos.
4. **Documentação de Cálculos**: Manter um registro detalhado para futuras referências.

## 7. Formulação Completa do Problema de Otimização

### 7.1 Objetivo
Minimizar os custos totais associados à produção, incluindo materiais, operações e penalidades.

### 7.2 Função Objetivo
\[
\begin{aligned}
\text{Min } Z = & \sum_{l \in L} \sum_{c \in C} \left( \frac{L_l \times P_m \times c}{1000} + \frac{P_l \times C_c \times c}{100} + \frac{L_l \times C_s \times c}{1000} + C_o \times c + \frac{A_d \times C_d \times c}{10\,000} \right) \\
& + \sum_{l \in L} P_l \times Y_l + P_s \times S
\end{aligned}
\]

### 7.3 Restrições

#### 7.3.1 Utilização da Área
\[
\frac{\text{Área Útil}}{\text{Área Total}} \geq 0,80
\]

#### 7.3.2 Superprodução
\[
\frac{\text{Produção} - \text{Demanda}}{\text{Demanda}} \leq 0,05
\]

#### 7.3.3 Número de Camadas
\[
c \leq 50 \quad \forall c \in C
\]

#### 7.3.4 Limites Operacionais
\[
\begin{cases}
L_l \leq 10\,000 \quad \forall l \in L \\
c \leq 30 \quad \forall c \in C \\
\text{Largura do Tecido} \leq 1\,500 \text{ mm}
\end{cases}
\]

#### 7.3.5 Definições das Variáveis Binárias e Contínuas
\[
Y_l \in \{0, 1\} \quad \forall l \in L
\]
\[
S \geq 0
\]
\[
c \in \mathbb{N} \quad \forall c \in C
\]




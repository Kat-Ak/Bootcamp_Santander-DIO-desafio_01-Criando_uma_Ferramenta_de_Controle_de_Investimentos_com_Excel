# Desafio 01 · Criando uma Ferramenta de Controle de Investimentos com Excel

# 🪙 FIINVEST🔹INTELIGÊNCIA & ALOCAÇÃO PATRIMONIAL com Excel

Este desafio faz parte do Bootcamp **Santander - Excel com IA e Claude** desenvolvido em parceria com a plataforma da [DIO](https://dio.me).

O projeto consiste no desenvolvimento do **FIINVEST**, um aplicativo de tela limpa voltado à simulação de juros compostos, análise de cenários multidimensionais e distribuição automatizada de capital no mercado de Fundos Imobiliários (FIIs), desenhado sob a perspectiva *beyond-tech* e de minimalismo estrutural.

---

## 🧬 Estrutura do Repositório | Índice Prático

🗂 Desafio_01-Criando_uma_Ferramenta_de_Controle_de_Investimentos_com_Excel
├── 📁 evidencias/             # Capturas de tela do sistema operando (Formatos PNG/JPG)
├── 📄 Simulador_FII_KatAk.xlsx # Arquivo executável da planilha (Gabarito prático de backoffice)
└── 📝 README.md               # Bastidores analíticos, documentação lógica e governança

---

## 🎯 Resumo do Enunciado

*   **Desafio:** Construir do zero uma ferramenta de simulação patrimonial focada em Fundos Imobiliários, conectando de forma fluida aportes mensais, horizontes temporais e perfis de risco.
*   **Escopo Principal:** Desenvolver uma interface integrada com "cara de aplicativo" que automatize cálculos matemáticos complexos e divida o capital de forma inteligente entre os principais segmentos imobiliários.
*   **Funcionalidades Centrais:**
    *   Módulo de simulação de patrimônio a médio prazo e geração de renda passiva.
    *   Módulo de projeção em cascata para cenários temporais de 2 a 30 anos.
    *   Sistema de alocação tática por perfil de risco integrado a chaves lógicas compostas.
    *   Uso rigoroso de intervalos nomeados para eliminar coordenadas abstratas e simplificar a auditoria.
*   **Resultado Esperado:** Uma solução corporativa segura e intuitiva, onde qualquer profissional ou investidor consiga ler cenários e tomar decisões sem necessidade de digitar ou compreender fórmulas.

---

## 📋 Cumprimento dos Requisitos do Desafio | Checklist

Abaixo está o mapeamento de conformidade com as metas obrigatórias e evoluções implementadas:

- [x] **Interface Premium "FIINVEST":** Ocultação de linhas de grade, cabeçalhos de linha/coluna e barra de fórmulas. Aplicação da paleta corporativa em Azul Petróleo Profundo (HEX `#0A2F41`) nos banners de fundo, utilizando o emoji de moeda (`🪙`) como âncora funcional e fonte em **Dourado Champagne (HEX `#FAE284`) em negrito** para contraste de alto padrão com a tipografia *Segoe UI*.
- [x] **Mapeamento de Linhas e Ergonomia:** Inclusão de linhas de respiro estruturais (linhas 2, 7, 14, 18 e 27) para delimitar as seções e evitar a sobrecarga cognitiva do usuário final.
- [x] **Cálculo de Patrimônio Líquido:** Integração da função financeira `VF` para prever o montante acumulado e cálculo indexado de dividendos mensais.
- [x] **Projeção de Cenários em Escala:** Estruturação da matriz temporal estendida do ano 2 ao ano 30 (linha 34), utilizando travamentos com referências absolutas (`$`).
- [x] **Configurações & Governança de Dados:** Implementação da diretriz de investimento de 30% do salário e substituição completa de coordenadas por Intervalos Nomeados.
- [x] **Divisão Automatizada por Perfil:** Criação de validação de dados por lista suspensa na célula `perfil_escolhido`, integrada ao motor de busca por chave combinada.

---

## ⚛️ Engenharia de Fórmulas e Processos por Trás do Sistema

### 1. Matriz Dinâmica de Risco via Chave Composta (`PROCV`)
Para fatiar o capital automaticamente na tabela de alocação de ativos (linhas 20 a 25) com base na escolha do investidor:
*   **Fórmula de Alocação (Coluna B, Linha 20):** `=PROCV(perfil_escolhido & "-" & A20; apoio!$A:$D; 4; FALSO)`
*   **Fórmula de Distribuição Monetária (Coluna C, Linha 20):** `=B20 * valor_investido`
*   **Fórmula do Totalizador de Aportes (Célula C26):** `=SOMA(C$20:C$25)`
*   *Explicação Non-Tech:* O Excel concatena o perfil selecionado (B16) e o nome do ativo (A20), gerando uma identificação única (ex: `Moderado-TIJOLO`). O comando `PROCV` localiza esse código na aba de `apoio`, extrai a taxa percentual e a multiplica pelo capital travado em `valor_investido` (B17). O fechamento exato do aporte de R$ 300,00 é auditado e consolidado pela soma segura na linha 26.

### 2. Projeção Temporal de Longo Prazo (`VF`)
Para estimar a evolução patrimonial ao longo das décadas no Simulador de Cenários (linhas 30 a 34):
*   **Fórmula de Projeção (Coluna B, Linha 30):** `=VF(taxa_mensal; A$30*12; aporte*-1)`
*   **Fórmula de Renda Passiva (Coluna C, Linha 30):** `=B$30 * rendimento_carteira`
*   *Explicação Non-Tech:* A função financeira calcula o crescimento dos juros compostos multiplicando os anos da coluna A por 12 meses, aplicando o intervalo nomeado `taxa_mensal` e o `aporte` original fixado. O ciclo se encerra na linha 34, projetando o cenário de 30 anos com precisão absoluta.

---

## 📊 Estratégia de Alocação (Refinamento da Grade de Ativos)

A matriz de distribuição de ativos foi estruturada de forma a equilibrar a segurança patrimonial e a flexibilidade de mercado, aplicando critérios claros de simplificação:
*   **Consolidação de Ativos:** Os segmentos foram integrados diretamente nas categorias principais de imóveis físicos, evitando pulverizações e otimizando a leitura visual para o usuário.
*   **Inclusão de Multiestratégia:** Adicionou-se a classe de Fundos Multiestratégia, conferindo ao simulador uma camada de alocação mais flexível, dinâmica e resiliente para os cenários de médio e longo prazo.

| Tipo de FII | Perfil Conservador | Perfil Moderado | Perfil Agressivo |
| :--- | :---: | :---: | :---: |
| **TIJOLO** (Imóveis Físicos) | 50% | 30% | 10% |
| **PAPEL** (Certificados/Crédito) | 25% | 30% | 40% |
| **HÍBRIDO** (Misto) | 5% | 8% | 8% |
| **FOF** (Fundo de Fundos) | 7% | 8% | 5% |
| **MULTIESTRATÉGIA** (Flexível) | 13% | 19% | 27% |
| **DESENVOLVIMENTO** (Projetos) | 0% | 5% | 10% |
| **TOTAL VERIFICADO** | **100%** | **100%** | **100%** |

---

## 📈 Inteligência de Negócio: Indicador "Bola de Neve" [Autossuficiência Patrimonial] (Linhas 36 e 38)

Foi implementado um cálculo tático que demonstra o **Ponto de Virada Financeira**. 
Afastando-se de jargões matemáticos complexos, o sistema conta com um módulo de tradução de cenários desenhado sob uma abordagem humanizada e de gamificação (UX), transformando dados brutos em feedbacks emocionais e preditivos em tempo real:
*   **Texto de Entrada (Célula A36):** `Preço da Cota do FII que você quer calcular ➔`
*   **Texto de Resultado (Célula A38):** `Resultado do seu efeito bola de neve ➔`
*   **Fórmula Lógica do Visor (Célula B38):** `=SE(C13>=B36; "🎉 PARABÉNS! O seu ganho mensal já compra + " & INT(C13/B36) & " cotas sem mexer no seu salário."; "⏳ QUASE LÁ! Ainda faltam R$ " & ARRED(B36-C13;2) & ". Continue assim! Logo o seu ganho mensal cobrirá o preço de uma cota.")`
*   *Abordagem de Usabilidade & UX:* O sistema cruza o dividendo real de R\$ 78,69 (`C13`) com o preço inserido livremente pelo usuário em `B36` (ex: R$ 10,00). Em caso de sucesso, celebra a conquista com a insígnia de comemoração festiva (`🎉`). Caso o montante seja menor, o motor lógico realiza uma subtração contábil automática e renderiza a expressão de impacto positivo "QUASE LÁ! Continue assim!" (`⏳`), apontando o valor residual exato e arredondado (`ARRED`) que separa o investidor do ponto de virada, estimulando a regularidade financeira de forma acolhedora.
O investidor informa o valor que pretende aportar e o sistema calcula de acordo com esse valor se o investidor atingiu a autossuficiência em cenários onde a renda recorrente ultrapassa o custo de uma cota.
*   **Resultado Prático:** Dependendo da regularidade e dos valores dos aportes, esse tipo de recorrência de reinvestimentos, gera autossuficiência de forma exponencial, sem necessidade de novos desembolsos de salário.

---

## 🧠 Inventário de Intervalos Nomeados Criados

As variáveis centrais foram batizadas na Caixa de Nome do Excel, forçando o "Preto Puro" (RGB 0,0,0) nos valores claros para contraste de alta nitidez e ativando a **Formatação Condicional Cromática** na célula de resposta (`B38` acende em Dourado Champagne para o sucesso e em Laranja Bronze para o processo em maturação):
*   `aporte` ➔ Vinculado à célula de entrada original do capital (`C9`).
*   `tempo_anos` ➔ Vinculado ao prazo selecionado, aplicando a máscara personalizada `0" anos"` (`C10`).
*   `taxa_mensal` ➔ Vinculado à taxa de rendimento estimada para a simulação (`C11`).
*   `rendimento_carteira` ➔ Vinculado à taxa padrão histórica de dividendos (`C5`).
*   `perfil_escolhido` ➔ Vinculado ao seletor de risco com formato de botão executivo azul suave (`B16`).
*   `valor_investido` ➔ Célula automatizada de espelhamento do capital para a tabela de FIIs (`B16`), protegida contra alterações para mitigar inconsistências operacionais.
*   `B36` ➔ Variável aberta para inserção livre do preço do ativo estudado pelo investidor.
*   `C13` **(Destaque de Ouro):** Célula de dividendos configurada com preenchimento **Dourado Suave (RGB 250,207,97)**, protegida e travada como o principal indicador de valor do negócio.

---

## 🛡️ Proteção & Governança de Processos

O sistema foi estruturado sob uma **política de segurança multicamadas**:
1.  **Aba APP:** Configuração de intervalos editáveis rígidos, mantendo apenas as 7 células de entrada desbloqueadas para digitação e interação do usuário, enquanto o motor de cálculo (fórmulas e visores) foi completamente congelado contra edições.
2.  **Aba apoio:** Configurada sob congelamento total (100% protegida), blindando a retaguarda, as tabelas de referência e as chaves compostas contra manipulações indevidas.
3.  **Auditabilidade:** O recurso opera sem senhas restritivas de abertura, assegurando que avaliadores e recrutadores possam auditar a engenharia das fórmulas de forma transparente.

---

## 🗺️ Conexões do Ecossistema

*   🔗 **Projeto Full:** [Acessar Pasta do Desafio](https://github.com)
*   🗃️ **HUB Principal:** [Retornar ao Portfólio de Projetos](https://github.com)
*   🆔 **Perfil Profissional:** [Visualizar Identidade GitHub](https://github.com)
*   🌌 **Presença Digital:** [Conectar via LinkedIn & Ecossistema Visual](https://github.com)

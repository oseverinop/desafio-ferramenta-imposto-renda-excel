# Desafio: Criando Um Organizador de Declaração de Imposto de Renda no Microsoft Excel

## Entendendo Desafio 
Agora é a sua hora de brilhar e construir um perfil de destaque na DIO! Explore todos os conceitos abordados até aqui, aplique os conhecimentos adquiridos nas aulas e documente sua experiência para demonstrar sua compreensão dos temas discutidos.

## Descrição do Desafio
Este projeto tem como objetivo criar uma ferramenta no Excel que ajude a organizar e reunir informações essenciais para a declaração de imposto de renda. A proposta é construir um agregador de dados no qual o usuário possa controlar suas entradas de maneira eficiente e validada, com menus de navegação, validações automáticas e funcionalidades extras, como links rápidos. A solução será completamente construída no Excel, com recursos que tornam a ferramenta robusta, mas com uma interface amigável e prática.

## Instrutor
- [Felipe Aguiar](https://www.linkedin.com/in/felipeaguiar-exe/)

## RELATÓRIO TÉCNICO DO PROCESSO DE CRIAÇÃO DA FERRAMENTA

### Introdução
Este relatório técnico detalha o processo de desenvolvimento de uma **planilha avançada no Microsoft Excel**, projetada para simplificar e organizar a coleta de informações necessárias para a **declaração anual de imposto de renda**. Ao longo deste documento, serão exploradas as etapas de construção, os recursos e conceitos do Excel empregados, e a importância de cada um para transformar uma simples grade de dados em uma **ferramenta interativa e profissional**. O objetivo final é apresentar um guia claro sobre como as funcionalidades do Excel podem ser combinadas para criar soluções eficientes e com uma experiência de usuário aprimorada.

---

### 1. Estrutura Inicial e Menu Básico

**Procedimento:** A etapa inicial focou na organização fundamental da planilha. As barras de título e fórmula do Excel foram ocultadas para uma visualização mais limpa. Em seguida, os botões do menu ("Titular", "Informes", "Notas") foram criados utilizando a ferramenta "Retângulo" do Excel. Cada botão foi configurado para não se mover ou redimensionar com as células, garantindo estabilidade visual.

**Conceitos e Recursos Abordados:**
* **Ocultar Elementos de Interface (Barra de Títulos e Fórmulas):** Recurso que permite remover elementos padrão do Excel da visualização.
* **Inserção e Formatação de Formas:** Utilização de objetos gráficos pré-definidos (retângulos) para criar elementos de interface.
* **Propriedades de Objeto ("Não mover ou dimensionar com células"):** Configuração que impede que um objeto gráfico altere sua posição ou tamanho quando as células da planilha são modificadas.

**Importância:** Essencial para estabelecer uma **base visual limpa e profissional**, transformando a planilha de uma grade de dados em uma interface mais amigável. A estabilidade dos botões do menu é crucial para a consistência do design.

---

### 2. Configuração Visual e Nomenclatura das Abas

**Procedimento:** A estética do menu foi aprimorada, alinhando os botões e alterando suas cores. As abas foram renomeadas para "Titular", "Informes" e "Notas" para refletir as seções do formulário. A visualização das linhas de grade do Excel também foi desativada para todas as abas.

**Conceitos e Recursos Abordados:**
* **Alinhamento e Distribuição de Objetos:** Ferramentas que permitem organizar elementos gráficos de forma simétrica e uniforme.
* **Cores e Estilos:** Aplicação de paletas de cores e preenchimentos para aprimorar o design visual.
* **Renomear Abas:** Modificação dos nomes das guias da planilha.
* **Ocultar Linhas de Grade:** Remoção das linhas que delimitam as células na visualização da planilha.

**Importância:** Contribui para a **identidade visual** e a **organização do conteúdo**, tornando a navegação mais intuitiva. A ocultação das linhas de grade proporciona uma aparência mais próxima de uma aplicação de software.

---

### 3. Funcionalidade do Menu e Duplicação de Abas

**Procedimento:** O menu lateral foi tornado funcional através da aplicação de hiperlinks. Cada botão foi vinculado à sua respectiva aba dentro do documento. Para agilizar o processo e manter a uniformidade, a aba "Titular" foi duplicada (usando `Ctrl` + arrastar a aba) para criar as abas "Informes" e "Notas". Após a duplicação, a cor de preenchimento dos botões do menu foi ajustada em cada aba para destacar a seção ativa. Um link externo (ícone do LinkedIn) também foi adicionado, demonstrando a capacidade de integrar links para fora da planilha.

**Conceitos e Recursos Abordados:**
* **Hiperlinks Internos ("Colocar neste documento"):** Permitem a navegação rápida e eficiente entre as abas da mesma pasta de trabalho.
* **Duplicação de Planilhas:** Criação de cópias idênticas de uma aba existente, incluindo seus objetos e formatações.
* **Feedback Visual (Destaque de Botões):** Técnica de design que informa ao usuário sobre o estado atual da interface (neste caso, qual aba está ativa).
* **Hiperlinks Externos ("Página Web ou Arquivo Existente"):** Permitem vincular a planilha a recursos fora do arquivo Excel, como websites.

**Importância:** Transforma o menu em um **sistema de navegação interativo**, fundamental para a usabilidade da ferramenta. A duplicação economiza tempo e assegura a consistência do layout em todas as seções.

---

### 4. Alinhamento Preciso de Elementos com VBA

**Procedimento:** Para garantir o alinhamento exato de elementos gráficos entre as abas, um código VBA (Visual Basic for Applications) foi utilizado. Primeiramente, o ícone do LinkedIn teve seu nome interno alterado no Painel de Seleção. Em seguida, via `Alt + F11`, um módulo VBA foi inserido, e o código foi colado para definir as coordenadas X e Y do ícone. O código foi executado em cada aba para posicionar o ícone de forma idêntica. Após o alinhamento, o módulo VBA foi removido para que a planilha pudesse ser salva no formato `.xlsx`.

**Conceitos e Recursos Abordados:**
* **VBA (Visual Basic for Applications):** Linguagem de programação que estende as capacidades do Excel, permitindo automação e controle preciso de objetos.
* **Editor VBA e Módulos:** Ambiente de desenvolvimento para escrever e armazenar códigos VBA.
* **Painel de Seleção:** Ferramenta para gerenciar e renomear objetos gráficos na planilha.
* **Posicionamento de Objetos via Coordenadas X e Y:** Controle programático da localização exata de elementos na planilha.

**Importância:** Soluciona o desafio de **alinhar perfeitamente elementos gráficos** entre abas, algo difícil de obter manualmente. O VBA proporciona um nível de precisão que aprimora a qualidade visual e a uniformidade da interface, mesmo em layouts complexos.

---

### 5. Criação do Formulário "Titular"

**Procedimento:** A aba "Titular" foi desenvolvida como um formulário para coletar dados pessoais. Campos como Nome, CPF, Nascimento, etc., foram inseridos e formatados com fonte padronizada e bordas discretas. Um título para a seção foi adicionado, e a validação de dados foi aplicada a campos de "sim ou não" para criar listas suspensas. Um botão "Próximo" foi criado, vinculado à próxima aba.

**Conceitos e Recursos Abordados:**
* **Estruturação de Formulários:** Organização de campos para entrada de dados.
* **Formatação de Texto e Bordas:** Aplicação de estilos para legibilidade e estética.
* **Validação de Dados (Lista Suspensa):** Recurso que restringe a entrada de dados a opções predefinidas, garantindo consistência.
* **Navegação Sequencial (Botão "Próximo"):** Criação de um fluxo de navegação linear entre as seções.

**Importância:** Esta etapa é crucial para a **coleta organizada e controlada de informações**. A validação de dados minimiza erros de entrada, e a estrutura de formulário melhora a experiência do usuário.

---

### 6. Formatações Numéricas Personalizadas

**Procedimento:** Os campos de CPF, CEP, Telefone e Celular receberam formatações numéricas personalizadas para exibir os dados no padrão brasileiro, mesmo que o usuário digite apenas os números. Para o e-mail, foi configurado um hiperlink que, ao ser clicado, abre o cliente de e-mail com um assunto predefinido.

**Conceitos e Recursos Abordados:**
* **Formatação de Números Personalizada (`Custom Number Format`):** Permite definir máscaras de exibição para números (ex: `000"."000"."000"-"00` para CPF).
* **Máscaras de Entrada:** Padrões de formatação que guiam a exibição dos dados.
* **Hiperlink de E-mail com Assunto Predefinido:** Configuração de um link de e-mail que inclui um campo de assunto pré-preenchido.

**Importância:** Aumenta a **legibilidade e a usabilidade** do formulário ao apresentar informações como CPF e telefones em um formato familiar. A personalização do link de e-mail otimiza a comunicação, tornando a ferramenta mais interativa.

---

### 7. Criação da Tela "Informes"

**Procedimento:** A aba "Informes" foi desenvolvida para registrar os informes bancários. O layout seguiu o padrão visual da aba "Titular". Foram criados campos para "Banco", "Valor Atual" e "Anexo". Uma nova aba oculta chamada "Tabelas" foi criada para armazenar uma lista de códigos e nomes de bancos, que foi usada na validação de dados do campo "Banco" (lista suspensa). Um totalizador de valores e botões de navegação "Anterior" e "Próximo" foram adicionados.

**Conceitos e Recursos Abordados:**
* **Consistência e Reutilização de Design:** Aplicação dos mesmos estilos e larguras de coluna das abas anteriores.
* **Abas de Apoio Ocultas:** Utilização de planilhas auxiliares para dados de referência, que são ocultadas da visualização do usuário.
* **Validação de Dados com Lista de Fonte Externa (Aba Oculta):** Restrição de entrada de dados com base em uma lista localizada em outra aba.
* **Mensagens de Entrada e Alertas de Erro:** Orientação e feedback ao usuário durante o preenchimento de campos.
* **Função SOMA:** Agregação de valores numéricos.

**Importância:** Expande a funcionalidade da planilha para **coleta de dados financeiros**, garantindo a precisão e padronização através da validação de dados. A organização com abas de apoio e a navegação clara aprimoram a robustez da ferramenta.

---

### 8. Criação da Tela "Notas"

**Procedimento:** A última aba principal, "Notas", foi criada para registrar entradas financeiras por categoria (Holerite, CNPJ, Freelance). O layout manteve a consistência visual das abas anteriores. Uma tabela foi configurada para registrar a "Data de Entrada" (formatada para Mês/Ano), "Categoria" (com lista suspensa) e "Valor" (formatado como moeda). O botão "Anterior" foi adicionado para navegar de volta à aba "Informes".

**Conceitos e Recursos Abordados:**
* **Organização de Dados em Tabela:** Estruturação de informações em formato tabular para clareza.
* **Formatação Personalizada de Data:** Exibição de datas em formatos específicos (ex: `MMMM/AAAA`).
* **Validação de Dados para Categorização:** Restrição de categorias para garantir consistência dos dados de entrada.
* **Navegação Direcional:** Manutenção de botões de navegação para um fluxo de usuário intuitivo.

**Importância:** Completa a estrutura de coleta de dados, permitindo o registro detalhado de diferentes tipos de receitas. A padronização e validação de dados nesta etapa são fundamentais para a **análise futura e a precisão das informações financeiras**.

---

### 9. Toques Finais e Proteção da Planilha

**Procedimento:** A interface foi finalizada ocultando-se a barra de fórmulas, a barra de títulos e configurando o modo de tela cheia. A principal medida foi a **proteção das planilhas**: as células destinadas à entrada de dados foram **desbloqueadas**, enquanto todas as outras (títulos, textos explicativos, áreas de cálculo, etc.) permaneceram bloqueadas. Em seguida, a proteção da planilha foi ativada, permitindo que o usuário apenas selecione e interaja com as células desbloqueadas e os hiperlinks.

**Conceitos e Recursos Abordados:**
* **Ajustes de UI (Modo Tela Cheia, Ocultar Barras):** Refinamento da aparência para simular um aplicativo.
* **Proteção de Planilha (`Protect Sheet`):** Recurso de segurança que impede alterações não autorizadas.
* **Células Bloqueadas/Desbloqueadas:** Propriedade de células que define se podem ser editadas quando a planilha está protegida.
* **Controle de Interação do Usuário:** Restrição de ações do usuário para garantir a integridade do layout e dos dados.

**Importância:** Esta etapa é crucial para a **robustez e profissionalismo** da ferramenta. A proteção da planilha garante a **integridade do design e dos dados**, evitando edições acidentais e direcionando o usuário apenas para os campos de entrada, transformando a planilha em um formulário interativo e seguro, ideal para portfólio ou uso prático.

---
### Conclusão

Em suma, a criação desta planilha de Excel demonstra o vasto potencial da ferramenta em ir além de suas funções básicas de cálculo e organização de dados. Ao integrar **conceitos de design de interface**, **automação via VBA**, **validação de dados** e **proteção de planilhas**, foi possível desenvolver uma aplicação robusta e intuitiva para a organização de informações fiscais. O resultado é uma ferramenta que não só simplifica a tarefa anual de reunir dados para o imposto de renda, mas também eleva a percepção de valor de uma planilha Excel, transformando-a em uma **solução profissional e altamente funcional**. Este projeto serve como um excelente exemplo de como a atenção aos detalhes e a aplicação estratégica de recursos podem gerar resultados impressionantes, impactando positivamente a usabilidade e a experiência do usuário.

---

> Este projeto foi desenvolvido como parte do bootcamp da DIO - Santander - Excel com Inteligência Artificial.

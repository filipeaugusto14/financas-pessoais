# Minhas Finanças — App de gestão de receitas e despesas

Web app em **Google Apps Script** sobre a planilha Google **"Finanças pessoais"**. Serve para:

1. **Registrar** receitas e despesas de forma rápida e amigável (celular ou desktop).
2. Ver um **relatório/dashboard** completo (KPIs, gráficos, rankings, indicadores), replicando o painel do Looker Studio com melhorias.

O app **mantém a aba `DADOS`** atualizada (valores com sinal, parcelas explodidas mês a mês com sufixo `(n/total)`), então o seu Looker Studio atual continua funcionando normalmente.

---

## Como a planilha é usada

| Aba | Papel | Colunas |
|-----|-------|---------|
| Lançamentos (entrada) | O que você digita | `DATA · VALOR (total) · DESCRIÇÃO · SUBCATEGORIA · CATEGORIA · TIPO · PARCELAS · ENVIADO PARA DADOS · ID` |
| `DADOS` | Consolidado que o relatório lê | `DATA · VALOR (com sinal) · DESCRIÇÃO · SUBCATEGORIA · CATEGORIA · TIPO · ID` |
| Taxonomia | Referência `SUBCATEGORIA → CATEGORIA → TIPO` | — |

> A coluna **`ID`** é adicionada automaticamente (pela função `setup()`) ao final das abas de Lançamentos e `DADOS`. É um campo extra ignorado pelo Looker; serve para o app editar/excluir com segurança.

O app **detecta as abas automaticamente** pela assinatura das colunas. Se algo não bater, rode `diagnostico()` (veja abaixo) e, se preciso, preencha os overrides no topo de `src/Config.gs`.

---

## Instalação (via clasp)

Pré-requisitos: **Node** e **clasp** já instalados (`clasp -v`).

```bash
# 1. Autenticar na sua conta Google
clasp login

# 2. Dentro da pasta do projeto, criar o projeto Apps Script (standalone)
cd financas-pessoais
clasp create --type standalone --title "Minhas Finanças" --rootDir src
#   -> isso gera o .clasp.json com o scriptId (fora do git)

# 3. Enviar o código
clasp push

# 4. Abrir o editor para configurar e implantar
clasp open
```

No editor do Apps Script:

1. Confirme o `SPREADSHEET_ID` no topo de `Config.gs` (já vem preenchido com a sua planilha).
2. Rode a função **`setup`** uma vez (autorize os acessos quando pedir). Ela verifica as abas e cria a coluna `ID`.
3. (Opcional) Rode **`diagnostico`** e veja no log (`Ctrl+Enter`) quais abas foram detectadas.
4. **Implantar › Nova implantação › App da Web**
   - *Executar como:* **Eu**
   - *Quem pode acessar:* **Somente eu**
5. Copie a URL do app e abra no celular/navegador.

Para atualizar depois de mudar o código: `clasp push` e (se mudou o comportamento) crie uma nova versão da implantação.

---

## Estrutura do código (`src/`)

| Arquivo | Responsabilidade |
|---------|------------------|
| `Config.gs` | ID da planilha, detecção de abas, `setup()`, `diagnostico()` |
| `Code.gs` | `doGet()` (serve o app), `include()` |
| `Lancamentos.gs` | Adicionar/listar/editar/excluir lançamentos |
| `Dados.gs` | Explosão de parcelas → aba `DADOS` |
| `Taxonomia.gs` | Categorias/subcategorias para o formulário |
| `Relatorio.gs` | Agregações do dashboard por período |
| `Index.html` · `Styles.html` · `App.html` | Interface (formulário + dashboard) |

---

## Notas

- **Parcelas:** ao lançar `R$ 300` em `3x`, a aba `DADOS` recebe 3 linhas mensais de `-100,00` com `(1/3) … (3/3)`. O arredondamento é em centavos e a diferença cai na última parcela (a soma bate com o total).
- **Sinal:** despesa entra negativa na `DADOS`, receita positiva.
- **Privacidade:** o app é publicado com acesso "Somente eu"; os dados nunca saem da sua conta Google.

# 🌾 Prescritor de Biológicos Syngenta (WebApp Apps Script)

Aplicação Web desenvolvida em Google Apps Script para recomendação e cálculo de dosagem de produtos biológicos em campo com exportação formatada para WhatsApp.

## 📊 Estrutura da Planilha Google Sheets (Banco de Dados)

O script exige uma planilha contendo uma aba exatamente com o nome **`Base_Produtos`**.

### Colunas esperadas (A até I):

| Coluna | Nome | Exemplo |
| :--- | :--- | :--- |
| **A** | ID | `1` |
| **B** | Cultura | `Soja` |
| **C** | Alvo | `Nematoides` |
| **D** | Produto | `Clariva PN` |
| **E** | Dose_Ha | `0.2` |
| **F** | Unidade_Dose | `L/ha` |
| **G** | Vol_Calda_Ha | `100` |
| **H** | Unidade_Calda | `L/ha` |
| **I** | Observacoes_Tecnicas | `Aplicar via sulco de plantio` |

---

## 🔒 Segurança e Chave de Acesso

* **PIN / Senha:** `7410`
* **Acesso Direto (Atalho para Usuários Autorizados):**
  Adicione `?chave=7410` ao final da URL da WebApp para desbloquear o aplicativo automaticamente sem exibir a tela de PIN:
  `https://script.google.com/macros/s/SUA_URL_AQUI/exec?chave=7410`

---

## 🚀 Como Sincronizar e Publicar no VS Code

### 1. Pré-requisitos
* [Node.js](https://nodejs.org/) instalado
* Instalador global do `clasp`:
  ```bash
  npm install -g @google/clasp
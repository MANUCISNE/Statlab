# StatLab — Análise Estatística Automática

Ferramenta web que usa IA para sugerir e executar análises estatísticas automaticamente a partir de planilhas Excel.

## 📁 Estrutura de arquivos (importante!)

Os arquivos baixados precisam ser organizados assim no GitHub:

```
statlab/
├── package.json          ← raiz
├── vite.config.js        ← raiz
├── index.html            ← raiz
├── README.md             ← raiz (este arquivo)
├── api/
│   └── claude.js         ← dentro da pasta "api"
└── src/
    ├── main.jsx          ← dentro da pasta "src"
    └── App.jsx           ← dentro da pasta "src"
```

⚠️ **Atenção aos nomes:** os arquivos `App.jsx`, `main.jsx` e `claude.js` precisam estar nas pastas certas (`src/` e `api/`), senão o app não funciona.

## 🚀 Deploy no Vercel (passo a passo)

### 1. Criar repositório no GitHub
- Acesse https://github.com/new
- Nome: `statlab` (ou outro)
- Marque **Public**
- Marque **Add a README file**
- Crie

### 2. Subir os arquivos
No GitHub, clique em **"Add file" → "Upload files"** e arraste os 6 arquivos. **Importante:** depois de arrastar, ajuste os caminhos clicando no nome de cada um:
- `App.jsx` → renomeie o caminho para `src/App.jsx`
- `main.jsx` → renomeie para `src/main.jsx`
- `claude.js` → renomeie para `api/claude.js`
- Os outros (`package.json`, `vite.config.js`, `index.html`) ficam na raiz

Commit changes.

### 3. Obter API key da Anthropic
- https://console.anthropic.com → criar conta
- "Plans & Billing" → adicionar ~5€ de crédito
- "API Keys" → "Create Key" → copiar (começa com `sk-ant-...`)

### 4. Deploy no Vercel
- https://vercel.com/new → "Import" o repositório
- **NÃO mude** as configurações de build (Vercel detecta Vite sozinho)
- Em **"Environment Variables"** adicione:
  - Name: `ANTHROPIC_API_KEY`
  - Value: sua chave
- Clique **Deploy**
- ~2 minutos depois, sua URL aparece (ex: `statlab-xyz.vercel.app`)

## ⚠️ Importante sobre uso científico

Os cálculos são feitos por IA (Claude), não por bibliotecas estatísticas tradicionais.

- ✅ **Bom para:** exploração inicial, ensino, hipótese-teste rápida
- ❌ **Não use para:** publicação científica formal. Para isso, valide os resultados em R, SPSS ou Python (scipy/statsmodels) antes de reportar.

## 📝 Licença

Uso pessoal livre.

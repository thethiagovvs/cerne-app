# Cerne — Dashboard Financeiro

Dashboard financeiro pessoal (React + Vite), com dados salvos no navegador
(`localStorage`) e backup automático rotativo no Dropbox.

## 1. Rodar localmente

```bash
npm install
npm run dev
```

Abre em `http://localhost:5173`.

## 2. Configurar o backup no Dropbox (opcional, mas recomendado)

O app já vem com toda a lógica pronta — falta só criar seu app no Dropbox e
colar a chave. Leva uns 5 minutos, é gratuito, não pede cartão de crédito.

1. Acesse **https://www.dropbox.com/developers/apps** e faça login com sua conta Dropbox.
2. Clique em **Create app**.
3. Escolha:
   - **Scoped access**
   - **App folder** (assim o app só enxerga a própria pasta dele no seu Drive, nunca o resto)
   - Dê um nome ao app, ex: `cerne-financeiro-thiago` (o nome precisa ser único no Dropbox)
4. Na aba **Permissions** do app criado, marque:
   - `files.metadata.write`
   - `files.metadata.read`
   - `files.content.write`
   - `files.content.read`
   Clique em **Submit** para salvar.
5. Na aba **Settings**:
   - Copie o **App key** (não o App secret — não precisamos dele).
   - Em **OAuth 2 → Redirect URIs**, adicione as URLs onde o app vai rodar:
     - `http://localhost:5173/` (para testar localmente)
     - a URL final do seu deploy, ex: `https://seu-app.vercel.app/` (adicione depois que fizer o deploy)
6. Abra `src/dropboxSync.js` neste projeto e troque:
   ```js
   export const DROPBOX_APP_KEY = 'COLE_AQUI_SEU_APP_KEY';
   ```
   pelo App key copiado.
7. Rode `npm run dev` de novo, vá em **Configurações → Backup na nuvem** e clique em **Conectar Dropbox**.

Como funciona depois de conectado:
- A cada alteração (lançamento, meta, etc.), o app espera alguns segundos e sobe
  um arquivo `backup-<data>.json` para a pasta do app no seu Dropbox.
- Ele mantém só os **5 backups mais recentes** — os mais antigos são apagados automaticamente.
- Se você abrir o app em outro navegador/aparelho sem dados salvos localmente, ele
  detecta que o Dropbox está conectado e restaura o backup mais recente sozinho.
- Também dá pra restaurar manualmente a qualquer momento em Configurações.

## 3. Publicar de graça

Qualquer uma das opções abaixo funciona bem e é gratuita para uso pessoal.

### Opção A — Vercel (recomendada)
1. Suba este projeto para um repositório no GitHub.
2. Em https://vercel.com, clique em **Add New → Project** e importe o repositório.
3. O Vercel detecta que é um projeto Vite automaticamente — não precisa configurar nada.
4. Depois do primeiro deploy, copie a URL gerada (ex: `https://cerne-thiago.vercel.app`)
   e adicione ela nos **Redirect URIs** do seu app no Dropbox (passo 5 acima).

### Opção B — Netlify (sem precisar de Git)
1. Rode `npm run build` localmente — isso gera a pasta `dist/`.
2. Em https://app.netlify.com, arraste a pasta `dist/` para a área de deploy.
3. Copie a URL gerada e adicione nos Redirect URIs do Dropbox, como acima.

## 4. Instalar como app no celular (PWA)

Depois de publicado, abra a URL no celular:
- **Android (Chrome):** menu (⋮) → "Adicionar à tela inicial".
- **iPhone (Safari):** botão de compartilhar → "Adicionar à Tela de Início".

O ícone abre em tela cheia, sem barra de navegador, como um app normal.

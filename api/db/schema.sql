-- ============================================================
-- Cerne — schema do banco (Fase 1: fundação)
-- ============================================================
-- Como rodar: cole esse arquivo inteiro no "Query" do painel do Vercel
-- Postgres (Storage → seu banco → Query), ou rode via psql apontando pra
-- string de conexão que o Vercel te dá. É seguro rodar mais de uma vez
-- (IF NOT EXISTS em tudo).

-- Cada linha é um usuário do Cerne. A identidade de verdade vem do Dropbox
-- (dropbox_account_id, algo como "dbid:AABBcc..." que a API do Dropbox
-- devolve depois do login) — nunca um e-mail ou senha escolhidos aqui.
CREATE TABLE IF NOT EXISTS users (
  id                 UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  dropbox_account_id TEXT UNIQUE NOT NULL,
  alias              TEXT UNIQUE NOT NULL,
  status             TEXT NOT NULL DEFAULT 'active' CHECK (status IN ('active', 'blocked')),
  role               TEXT NOT NULL DEFAULT 'user' CHECK (role IN ('user', 'admin')),
  -- Tokens do Dropbox deste usuário, guardados aqui pra o backend conseguir
  -- sincronizar sem precisar do usuário logado o tempo todo. Nunca saem
  -- daqui pro navegador — só o backend os lê (ver Fase 2).
  dropbox_refresh_token TEXT,
  dropbox_access_token  TEXT,
  dropbox_token_expires_at TIMESTAMPTZ,
  created_at         TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at         TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Alias tem que ser único, mas a busca por ele (na hora do login, pra achar
-- o usuário) deve ignorar maiúsculas/minúsculas — "Thiago" e "thiago" não
-- podem ser dois usuários diferentes.
CREATE UNIQUE INDEX IF NOT EXISTS users_alias_lower_idx ON users (lower(alias));

-- Uma linha por sessão ativa. O cookie do navegador guarda só o "id" (um
-- token aleatório, não o id do usuário) — é o que prova que aquele
-- navegador está autenticado, sem precisar reconsultar o Dropbox a cada
-- clique.
CREATE TABLE IF NOT EXISTS sessions (
  id         TEXT PRIMARY KEY,
  user_id    UUID NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  expires_at TIMESTAMPTZ NOT NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS sessions_user_id_idx ON sessions (user_id);

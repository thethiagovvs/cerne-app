// Conexão com o Vercel Postgres, compartilhada por todas as rotas em /api.
// A variável de ambiente POSTGRES_URL é preenchida automaticamente pela
// Vercel quando você linka um banco Postgres ao projeto — não precisa
// configurar nada manualmente além de criar o banco pelo painel.
import { sql } from '@vercel/postgres';

export { sql };

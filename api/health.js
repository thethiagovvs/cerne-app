// GET /api/health — confirma que a conexão com o banco está funcionando e
// que a tabela "users" existe (ou seja, que o schema.sql já foi rodado).
// Só serve pra validar a Fase 1; não expõe nenhum dado de usuário.
import { sql } from './lib/db.js';

// Variáveis de ambiente que a Vercel costuma criar quando um Postgres é
// linkado ao projeto — o nome exato mudou algumas vezes ao longo do tempo
// conforme a Vercel evoluiu essa integração, por isso conferimos todas.
// Só reportamos SE existem, nunca o valor — não vaza segredo nenhum.
const EXPECTED_ENV_VARS = ['POSTGRES_URL', 'POSTGRES_PRISMA_URL', 'POSTGRES_URL_NON_POOLING', 'DATABASE_URL'];

export default async function handler(req, res) {
  const envStatus = Object.fromEntries(EXPECTED_ENV_VARS.map((name) => [name, !!process.env[name]]));
  const anyEnvPresent = Object.values(envStatus).some(Boolean);

  if (!anyEnvPresent) {
    res.status(500).json({
      status: 'error',
      message: 'Nenhuma variável de ambiente de banco foi encontrada nesta função. O Postgres provavelmente não está linkado a ESTE projeto específico na Vercel, ou o projeto não foi reimplantado (redeploy) depois de linkar o banco — variáveis novas só valem a partir do próximo deploy.',
      env_vars_found: envStatus,
    });
    return;
  }

  try {
    const result = await sql`SELECT count(*)::int AS user_count FROM users;`;
    res.status(200).json({
      status: 'ok',
      database: 'connected',
      users_table: 'exists',
      user_count: result.rows[0].user_count,
      env_vars_found: envStatus,
    });
  } catch (err) {
    res.status(500).json({
      status: 'error',
      message: 'Uma variável de ambiente do banco existe, mas a conexão ou a consulta falhou. Se a tabela "users" não existir ainda, rode o schema.sql. Se o erro for de conexão, confira se o banco não foi pausado/removido no painel da Vercel.',
      env_vars_found: envStatus,
      detail: err.message,
    });
  }
}

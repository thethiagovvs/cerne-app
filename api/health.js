// GET /api/health — confirma que a conexão com o banco está funcionando e
// que a tabela "users" existe (ou seja, que o schema.sql já foi rodado).
// Só serve pra validar a Fase 1; não expõe nenhum dado de usuário.
import { sql } from './lib/db.js';

export default async function handler(req, res) {
  try {
    const result = await sql`SELECT count(*)::int AS user_count FROM users;`;
    res.status(200).json({
      status: 'ok',
      database: 'connected',
      users_table: 'exists',
      user_count: result.rows[0].user_count,
    });
  } catch (err) {
    res.status(500).json({
      status: 'error',
      message: 'Banco não respondeu como esperado. Confira se o Postgres foi linkado ao projeto na Vercel e se o schema.sql já foi rodado.',
      detail: err.message,
    });
  }
}

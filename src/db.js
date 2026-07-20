import pg from "pg";
const { Pool } = pg;

console.log('DATABASE_URL:', process.env.DATABASE_URL ? process.env.DATABASE_URL.substring(0, 30) + '...' : 'NOT SET');

const connectionConfig = process.env.DATABASE_URL && !process.env.DATABASE_URL.includes('base')
  ? { connectionString: process.env.DATABASE_URL, ssl: { rejectUnauthorized: false } }
  : {
      host: process.env.PGHOST,
      port: process.env.PGPORT || 5432,
      database: process.env.PGDATABASE || process.env.POSTGRES_DB,
      user: process.env.PGUSER || process.env.POSTGRES_USER,
      password: process.env.PGPASSWORD || process.env.POSTGRES_PASSWORD,
      ssl: { rejectUnauthorized: false }
    };

const pool = new Pool(connectionConfig);

export async function query(text, params) {
  const client = await pool.connect();
  try {
    return await client.query(text, params);
  } finally {
    client.release();
  }
}

export async function initDB() {
  await query(`
    CREATE TABLE IF NOT EXISTS candidates (
      id SERIAL PRIMARY KEY,
      timestamp TIMESTAMPTZ DEFAULT NOW(),
      telegram_id BIGINT,
      telegram_username TEXT DEFAULT '',
      full_name TEXT,
      phone TEXT,
      email TEXT,
      city TEXT,
      degree TEXT,
      field_of_study TEXT,
      languages TEXT,
      is_intern TEXT,
      internship_mentor TEXT,
      internship_phone TEXT,
      experience TEXT,
      interests TEXT,
      workplace_pref TEXT,
      timing TEXT,
      availability TEXT,
      cv TEXT,
      motivation TEXT,
      has_references TEXT,
      "references" TEXT,
      political_side TEXT,
      has_license TEXT,
      english_level TEXT,
      irregular_hours TEXT,
      declaration TEXT,
      status TEXT DEFAULT 'active',
      created_at TIMESTAMPTZ DEFAULT NOW(),
      updated_at TIMESTAMPTZ DEFAULT NOW()
    )
  `);

  await query(`
    CREATE TABLE IF NOT EXISTS employers (
      id SERIAL PRIMARY KEY,
      timestamp TIMESTAMPTZ DEFAULT NOW(),
      telegram_id BIGINT,
      telegram_username TEXT DEFAULT '',
      org_type TEXT,
      contact_name TEXT,
      phone TEXT,
      email TEXT,
      fields TEXT,
      timing TEXT,
      availability TEXT,
      experience_importance TEXT,
      notes TEXT,
      political_side TEXT,
      requires_license TEXT,
      english_required TEXT,
      irregular_hours TEXT,
      declaration TEXT,
      status TEXT DEFAULT 'active',
      created_at TIMESTAMPTZ DEFAULT NOW(),
      updated_at TIMESTAMPTZ DEFAULT NOW()
    )
  `);

  await query(`
    CREATE TABLE IF NOT EXISTS approved_phones (
      phone TEXT PRIMARY KEY
    )
  `);

  await query(`
    CREATE TABLE IF NOT EXISTS matches (
      id SERIAL PRIMARY KEY,
      candidate_id BIGINT,
      employer_id BIGINT,
      candidate_name TEXT,
      employer_name TEXT,
      status TEXT DEFAULT 'active',
      matched_at TIMESTAMPTZ DEFAULT NOW(),
      closed_at TIMESTAMPTZ,
      UNIQUE(candidate_id, employer_id)
    )
  `);

  await query(`
    CREATE TABLE IF NOT EXISTS recommendations (
      id SERIAL PRIMARY KEY,
      candidate_id BIGINT UNIQUE,
      text TEXT,
      recommender_name TEXT,
      created_at TIMESTAMPTZ DEFAULT NOW()
    )
  `);

  await query(`
    CREATE TABLE IF NOT EXISTS access_requests (
      id SERIAL PRIMARY KEY,
      timestamp TIMESTAMPTZ DEFAULT NOW(),
      telegram_id BIGINT,
      telegram_username TEXT DEFAULT '',
      phone TEXT,
      full_name TEXT,
      role TEXT,
      job_search TEXT,
      heard_from TEXT,
      status TEXT DEFAULT 'pending',
      approved_at TIMESTAMPTZ,
      denied_at TIMESTAMPTZ
    )
  `);

  console.log("✅ DB tables initialized");
}

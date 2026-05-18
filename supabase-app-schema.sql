-- ================================================
-- Crecensa App Schema (ejecutar en SQL Editor)
-- Tablas optimizadas para el dashboard
-- ================================================

CREATE TABLE IF NOT EXISTS clientes (
  id BIGINT PRIMARY KEY,
  nom TEXT NOT NULL DEFAULT '',
  tel TEXT NOT NULL DEFAULT '',
  dpi TEXT NOT NULL DEFAULT '',
  vend TEXT NOT NULL DEFAULT '',
  est TEXT NOT NULL DEFAULT 'activo',
  nota TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS terrenos (
  id BIGINT PRIMARY KEY,
  sec TEXT NOT NULL DEFAULT '',
  lot TEXT NOT NULL DEFAULT '',
  pre NUMERIC NOT NULL DEFAULT 0,
  sal NUMERIC NOT NULL DEFAULT 0,
  est TEXT NOT NULL DEFAULT 'disponible',
  cli_id BIGINT,
  vend TEXT NOT NULL DEFAULT '',
  are TEXT NOT NULL DEFAULT '',
  proj TEXT NOT NULL DEFAULT '',
  nota TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS recibos (
  id BIGINT PRIMARY KEY,
  fec TEXT NOT NULL DEFAULT '',
  fec2 TEXT NOT NULL DEFAULT '',
  num TEXT NOT NULL DEFAULT '',
  cli_id BIGINT,
  ter_id BIGINT,
  mon NUMERIC NOT NULL DEFAULT 0,
  bol TEXT NOT NULL DEFAULT '',
  vend TEXT NOT NULL DEFAULT '',
  nota TEXT NOT NULL DEFAULT ''
);

-- Row Level Security (requerido para publishable key)
ALTER TABLE clientes ENABLE ROW LEVEL SECURITY;
ALTER TABLE terrenos ENABLE ROW LEVEL SECURITY;
ALTER TABLE recibos  ENABLE ROW LEVEL SECURITY;

CREATE POLICY "allow_all_clientes" ON clientes FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "allow_all_terrenos" ON terrenos FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "allow_all_recibos"  ON recibos  FOR ALL USING (true) WITH CHECK (true);

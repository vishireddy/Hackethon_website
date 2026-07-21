CREATE TABLE IF NOT EXISTS admins (
  id SERIAL PRIMARY KEY,
  username VARCHAR(120) NOT NULL UNIQUE,
  password_hash VARCHAR(255) NOT NULL,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS participants (
  id SERIAL PRIMARY KEY,
  team_name VARCHAR(255) NOT NULL,
  team_leader_name VARCHAR(255) NOT NULL,
  email_address VARCHAR(255) NOT NULL,
  mobile_number VARCHAR(20) NOT NULL UNIQUE,
  password_hash VARCHAR(255) NOT NULL,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS problems (
  id SERIAL PRIMARY KEY,
  problem_code VARCHAR(20) NULL,
  title VARCHAR(255) NOT NULL,
  description TEXT NOT NULL,
  persona TEXT NULL,
  problem_statement TEXT NULL,
  solution_idea TEXT NULL,
  why_selected TEXT NULL,
  key_technologies TEXT NULL,
  domain VARCHAR(120) NOT NULL,
  is_active SMALLINT NOT NULL DEFAULT 1,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_problems_domain ON problems (domain);
CREATE INDEX IF NOT EXISTS idx_problems_active ON problems (is_active);

CREATE TABLE IF NOT EXISTS problem_selections (
  id SERIAL PRIMARY KEY,
  participant_id INT NULL,
  problem_id INT NOT NULL,
  team_name VARCHAR(255) NOT NULL,
  team_leader_name VARCHAR(255) NOT NULL,
  email_address VARCHAR(255) NOT NULL,
  mobile_number VARCHAR(20) NOT NULL UNIQUE,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT fk_problem_selections_participant
    FOREIGN KEY (participant_id) REFERENCES participants(id)
    ON DELETE SET NULL,
  CONSTRAINT fk_problem_selections_problem
    FOREIGN KEY (problem_id) REFERENCES problems(id)
    ON DELETE CASCADE,
  CONSTRAINT uq_problem_selections_participant UNIQUE (participant_id)
);

CREATE INDEX IF NOT EXISTS idx_problem_selections_problem ON problem_selections (problem_id);
CREATE INDEX IF NOT EXISTS idx_problem_selections_created ON problem_selections (created_at);

CREATE TABLE IF NOT EXISTS settings (
  setting_key VARCHAR(120) PRIMARY KEY,
  setting_value TEXT NULL,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

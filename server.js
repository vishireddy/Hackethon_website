require("dotenv").config();

const express = require("express");
const path = require("path");
const fs = require("fs/promises");
const bcrypt = require("bcryptjs");
const XLSX = require("xlsx");
const { Pool } = require("pg");
const session = require("express-session");

const pgPool = new Pool({ connectionString: process.env.DATABASE_URL });
const pool = {
  async query(sql, params = []) {
    let i = 1;
    let pgSql = sql.replace(/\?/g, () => `$${i++}`);
    pgSql = pgSql.replace(/\bAS\s+([a-zA-Z0-9_]+)\b/g, 'AS "$1"');
    const res = await pgPool.query(pgSql, params);
    const rows = res.rows;
    rows.insertId = rows.length > 0 ? rows[0].id : null;
    rows.affectedRows = res.rowCount;
    return [rows];
  },
  async getConnection() {
    const client = await pgPool.connect();
    return {
      async query(sql, params = []) {
        let i = 1;
        let pgSql = sql.replace(/\?/g, () => `$${i++}`);
        pgSql = pgSql.replace(/\bAS\s+([a-zA-Z0-9_]+)\b/g, 'AS "$1"');
        const res = await client.query(pgSql, params);
        const rows = res.rows;
        rows.insertId = rows.length > 0 ? rows[0].id : null;
        rows.affectedRows = res.rowCount;
        return [rows];
      },
      async beginTransaction() { await client.query('BEGIN'); },
      async commit() { await client.query('COMMIT'); },
      async rollback() { await client.query('ROLLBACK'); },
      release() { client.release(); }
    };
  }
};

const app = express();
const PORT = Number(process.env.PORT || 3000);
const TEAM_LIMIT = process.env.TEAM_LIMIT || 8;
const isProduction = process.env.NODE_ENV === "production";
const ASSET_VERSION = "2026-04-29-02";
const DEFAULT_DOMAINS = [
  "AI",
  "FinTech",
  "Healthcare",
  "EdTech",
  "Cybersecurity",
  "PropTech",
  "Logistics",
  "LegalTech",
  "ClimaTech",
  "HRTech",
  "InsurTech",
  "Mental Health",
  "GovTech",
  "RetailTech",
  "FoodTech",
];

const PARTNERS = {
  Ecosystem: [
    { name: "t-hub", logo: "/images/ecosystem-partners/thubhyd_logo.jpeg" },
    { name: "astir ventures", logo: "/images/ecosystem-partners/astir_ventures_logo.jpeg" },
    { name: "Hyderabad Angels", logo: "/images/ecosystem-partners/hyderabad_angels_ventures_logo.jpeg" },
    { name: "DALLAS VENTURE CAPITAL", logo: "/images/ecosystem-partners/dallas_venture_capital_logo.jpeg" }
  ],
  Technology: [
    { name: "rava", logo: "/images/technology-partners/rava_ai_logo.jpeg" },
    { name: "OSEN", logo: "/images/technology-partners/1759527371334.jpeg" },
    { name: "mem0", logo: "/images/technology-partners/mem0_logo.jpeg" }
  ],
  Resource: [
    { name: "PROMPT TECHIES", logo: "/images/prompt-techies.png" }
  ],
  Community: [
    { name: "CodeQuestors", logo: "/images/community-partners/codequesters_logo.jpeg" },
    { name: "DRAPER", logo: "/images/community-partners/draper_startup_house_hyderabad_logo.jpeg" },
    { name: "NEXTGEN NEXUS", logo: "/images/community-partners/nextgen_nexus_logo.jpeg" },
    { name: "GROUND ZERO", logo: "/images/community-partners/groundzerocommunitie_logo.jpeg" },
    { name: "LINKEDINSPIRE", logo: "/images/community-partners/1757608545924.jpeg" },
    { name: "SCCSE", logo: "/images/community-partners/so_called_cs_engineers_logo.jpeg" }
  ],
  Media: ["SHOOTGRAM"]
};

const DEFAULT_PROBLEM_CATALOG = [];

// pool already defined above

function normalizeProblem(problem) {
  const problemCode = cleanOptionalText(problem.problemCode);
  const description = String(problem.description || "").trim();
  const persona = cleanOptionalText(problem.persona);
  const problemStatement = cleanOptionalText(problem.problemStatement);
  const solutionIdea = cleanOptionalText(problem.solutionIdea);
  const whySelected = cleanOptionalText(problem.whySelected);
  const keyTechnologyList = splitMultilineList(problem.keyTechnologies);
  const primaryStatement = problemStatement || description;

  return {
    ...problem,
    problemCode,
    description,
    persona,
    problemStatement: primaryStatement,
    solutionIdea,
    whySelected,
    keyTechnologyList,
    hasOverview: Boolean(problemStatement && problemStatement !== description),
    teamCount: Number(problem.teamCount || 0),
    availableSlots: Math.max(TEAM_LIMIT - Number(problem.teamCount || 0), 0),
    isFull: Number(problem.teamCount || 0) >= TEAM_LIMIT,
    searchText: `${problemCode || ""} ${problem.title} ${description} ${problem.domain} ${persona || ""} ${primaryStatement} ${solutionIdea || ""} ${whySelected || ""} ${keyTechnologyList.join(" ")}`
      .toLowerCase()
      .trim(),
  };
}

function groupProblemsByDomain(problems) {
  return problems.reduce((groups, problem) => {
    if (!groups[problem.domain]) {
      groups[problem.domain] = [];
    }
    groups[problem.domain].push(problem);
    return groups;
  }, {});
}

function setFlash(req, type, message) {
  req.session.flash = { type, message };
}

function consumeFlash(req) {
  const flash = req.session.flash || null;
  delete req.session.flash;
  return flash;
}

function requireAdmin(req, res, next) {
  if (!req.session.admin) {
    setFlash(req, "error", "Please sign in as admin to continue.");
    return res.redirect("/admin/login");
  }

  return next();
}

function requireParticipant(req, res, next) {
  if (!req.session.participant) {
    setFlash(req, "error", "Please sign up or log in as a participant first.");
    return res.redirect("/");
  }

  return next();
}

function validMobileNumber(number) {
  return /^[6-9]\d{9}$/.test(String(number || "").trim());
}

function validPassword(password) {
  return String(password || "").trim().length >= 6;
}

function validEmailAddress(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || "").trim());
}

function cleanOptionalText(value) {
  const normalized = String(value || "").trim();
  return normalized ? normalized : null;
}

function splitMultilineList(value) {
  return String(value || "")
    .split(/\r?\n/)
    .map((item) => item.trim().replace(/^[•*-]\s*/, ""))
    .filter(Boolean);
}

function escapeCsvValue(value) {
  const stringValue = String(value ?? "");
  if (/[",\n]/.test(stringValue)) {
    return `"${stringValue.replace(/"/g, '""')}"`;
  }
  return stringValue;
}

function formatSelectionTimestamp(value) {
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? "" : date.toISOString();
}

async function fetchAllSelections() {
  const [rows] = await pool.query(
    `SELECT
      ps.id,
      ps.participant_id AS participantId,
      ps.team_name AS teamName,
      COALESCE(NULLIF(ps.email_address, ''), ps.team_leader_name) AS emailAddress,
      ps.mobile_number AS mobileNumber,
      ps.created_at AS createdAt,
      p.id AS problemId,
      p.problem_code AS problemCode,
      p.title AS problemTitle,
      p.domain AS problemDomain
    FROM problem_selections ps
    INNER JOIN problems p ON p.id = ps.problem_id
    ORDER BY ps.created_at DESC`
  );

  return rows.map((row) => ({
    ...row,
    createdAtFormatted: formatSelectionTimestamp(row.createdAt),
  }));
}

async function ensureSettingsTable() {
  const defaultSettings = [
    { key: "whatsapp_url", value: "#" },
    { key: "linkedin_url", value: "#" },
    { key: "instagram_url", value: "#" },
  ];
  for (const { key, value } of defaultSettings) {
    try {
      await pool.query("INSERT INTO settings (setting_key, setting_value) VALUES (?, ?) ON CONFLICT (setting_key) DO NOTHING", [key, value]);
      await pool.query("UPDATE settings SET setting_value = ? WHERE setting_key = ? AND (setting_value = '' OR setting_value IS NULL)", [value, key]);
    } catch (error) {
      console.warn("Failed to insert setting", key, error.message);
    }
  }
}

async function fetchSettings() {
  try {
    const [rows] = await pool.query('SELECT setting_key, setting_value FROM settings');
    const settings = {};
    for (const row of rows) {
      settings[row.setting_key] = row.setting_value;
    }
    return settings;
  } catch (error) {
    return {};
  }
}

async function initializeDatabase() {
  const schemaPath = path.join(__dirname, "db", "schema.sql");
  const schema = await fs.readFile(schemaPath, "utf8");
  const statements = schema
    .split(";")
    .map((statement) => statement.trim())
    .filter(Boolean);

  for (const statement of statements) {
    await pool.query(statement);
  }
}

async function ensureProblemBriefColumns() {
  const problemColumns = [
    ["problem_code", "VARCHAR(20) NULL"],
    ["persona", "TEXT NULL"],
    ["problem_statement", "TEXT NULL"],
    ["solution_idea", "TEXT NULL"],
    ["why_selected", "TEXT NULL"],
    ["key_technologies", "TEXT NULL"],
  ];

  for (const [name, definition] of problemColumns) {
    try {
      await pool.query(`ALTER TABLE problems ADD COLUMN ${name} ${definition}`);
    } catch (error) {
      if (error.code !== "42701") {
        throw error;
      }
    }
  }
}

async function ensureParticipantSelectionSupport() {
  try {
    await pool.query(
      "ALTER TABLE problem_selections ADD COLUMN participant_id INT NULL"
    );
  } catch (error) {
    if (error.code !== "42701") {
      throw error;
    }
  }

  try {
    await pool.query(
      "ALTER TABLE problem_selections ADD CONSTRAINT fk_problem_selections_participant FOREIGN KEY (participant_id) REFERENCES participants(id) ON DELETE SET NULL"
    );
  } catch (error) {
    if (error.code !== "42P07" && error.code !== "42710") {
      throw error;
    }
  }

  try {
    await pool.query(
      "CREATE UNIQUE INDEX uq_problem_selections_participant ON problem_selections (participant_id)"
    );
  } catch (error) {
    if (error.code !== "42P07") {
      throw error;
    }
  }
}

async function ensureParticipantEmailColumns() {
  try {
    await pool.query(
      "ALTER TABLE participants ADD COLUMN email_address VARCHAR(255) NULL"
    );
  } catch (error) {
    if (error.code !== "42701") {
      throw error;
    }
  }

  try {
    await pool.query(
      "ALTER TABLE problem_selections ADD COLUMN email_address VARCHAR(255) NULL"
    );
  } catch (error) {
    if (error.code !== "42701") {
      throw error;
    }
  }

  await pool.query(
    `UPDATE participants
    SET email_address = team_leader_name
    WHERE (email_address IS NULL OR email_address = '')
      AND team_leader_name IS NOT NULL
      AND team_leader_name <> ''`
  );

  await pool.query(
    `UPDATE problem_selections
    SET email_address = team_leader_name
    WHERE (email_address IS NULL OR email_address = '')
      AND team_leader_name IS NOT NULL
      AND team_leader_name <> ''`
  );
}

async function syncDefaultProblemCatalog() {
  for (const item of DEFAULT_PROBLEM_CATALOG) {
    const [byCode] = await pool.query(
      "SELECT id FROM problems WHERE problem_code = ? LIMIT 1",
      [item.problemCode]
    );

    if (byCode.length > 0) {
      continue;
    }

    const [byTitle] = await pool.query(
      "SELECT id FROM problems WHERE title = ? LIMIT 1",
      [item.title]
    );

    if (byTitle.length > 0) {
      await pool.query(
        `UPDATE problems
        SET
          problem_code = COALESCE(problem_code, ?),
          domain = COALESCE(NULLIF(domain, ''), ?),
          description = COALESCE(NULLIF(description, ''), ?),
          persona = COALESCE(NULLIF(persona, ''), ?),
          problem_statement = COALESCE(NULLIF(problem_statement, ''), ?),
          solution_idea = COALESCE(NULLIF(solution_idea, ''), ?),
          why_selected = COALESCE(NULLIF(why_selected, ''), ?),
          key_technologies = COALESCE(NULLIF(key_technologies, ''), ?)
        WHERE id = ?`,
        [
          item.problemCode,
          item.domain,
          item.description,
          item.persona || null,
          item.problemStatement,
          item.solutionIdea || null,
          item.whySelected || null,
          item.keyTechnologies || null,
          byTitle[0].id,
        ]
      );
      continue;
    }

    await pool.query(
      `INSERT INTO problems
        (
          problem_code,
          title,
          domain,
          description,
          persona,
          problem_statement,
          solution_idea,
          why_selected,
          key_technologies
        )
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [
        item.problemCode,
        item.title,
        item.domain,
        item.description,
        item.persona || null,
        item.problemStatement,
        item.solutionIdea || null,
        item.whySelected || null,
        item.keyTechnologies || null,
      ]
    );
  }
}

async function ensureAdminAccount() {
  const username = (process.env.ADMIN_USERNAME || "admin").trim();
  const password = process.env.ADMIN_PASSWORD || "ChangeMe123!";

  const [admins] = await pool.query("SELECT id FROM admins WHERE username = ? LIMIT 1", [username]);
  if (admins.length > 0) {
    return;
  }

  const passwordHash = await bcrypt.hash(password, 12);
  await pool.query("INSERT INTO admins (username, password_hash) VALUES (?, ?)", [username, passwordHash]);
  console.log(`Bootstrapped admin account for username "${username}".`);
}

async function fetchParticipantSelection(participantId) {
  if (!participantId) {
    return null;
  }

  const [rows] = await pool.query(
    `SELECT
      ps.id,
      ps.created_at AS createdAt,
      p.id AS problemId,
      p.problem_code AS problemCode,
      p.title,
      p.domain,
      p.description,
      p.problem_statement AS problemStatement
    FROM problem_selections ps
    INNER JOIN problems p ON p.id = ps.problem_id
    WHERE ps.participant_id = ?
    LIMIT 1`,
    [participantId]
  );

  if (rows.length === 0) {
    return null;
  }

  return {
    ...rows[0],
    problemStatement: cleanOptionalText(rows[0].problemStatement) || rows[0].description,
    createdAtFormatted: formatSelectionTimestamp(rows[0].createdAt),
  };
}

async function fetchPublicData(participantId) {
  const [problemsRaw] = await pool.query(
    `SELECT
      p.id,
      p.problem_code AS problemCode,
      p.title,
      p.description,
      p.persona AS persona,
      p.problem_statement AS problemStatement,
      p.solution_idea AS solutionIdea,
      p.why_selected AS whySelected,
      p.key_technologies AS keyTechnologies,
      p.domain,
      p.created_at AS createdAt,
      COUNT(ps.id) AS teamCount
    FROM problems p
    LEFT JOIN problem_selections ps ON ps.problem_id = p.id
    WHERE p.is_active = 1
    GROUP BY
      p.id,
      p.title,
      p.description,
      p.problem_code,
      p.persona,
      p.problem_statement,
      p.solution_idea,
      p.why_selected,
      p.key_technologies,
      p.domain,
      p.created_at
    ORDER BY p.domain ASC, p.problem_code ASC, p.created_at DESC`
  );

  const [statsRows] = await pool.query(
    `SELECT
      (SELECT COUNT(*) FROM problems WHERE is_active = 1) AS problemCount,
      (SELECT COUNT(*) FROM problem_selections) AS selectionCount`
  );

  const problems = problemsRaw.map(normalizeProblem);
  const domains = Array.from(new Set([...DEFAULT_DOMAINS, ...problems.map((problem) => problem.domain)])).sort();
  const participantSelection = await fetchParticipantSelection(participantId);
  return {
    problems,
    problemGroups: groupProblemsByDomain(problems),
    domains,
    participantSelection,
    settings: await fetchSettings(),
    stats: {
      problemCount: Number(statsRows[0].problemCount || 0),
      selectionCount: Number(statsRows[0].selectionCount || 0),
    },
  };
}

async function fetchAdminData() {
  const [problemsRaw] = await pool.query(
    `SELECT
      p.id,
      p.problem_code AS problemCode,
      p.title,
      p.description,
      p.persona AS persona,
      p.problem_statement AS problemStatement,
      p.solution_idea AS solutionIdea,
      p.why_selected AS whySelected,
      p.key_technologies AS keyTechnologies,
      p.domain,
      p.is_active AS isActive,
      p.created_at AS createdAt,
      p.updated_at AS updatedAt,
      COUNT(ps.id) AS teamCount
    FROM problems p
    LEFT JOIN problem_selections ps ON ps.problem_id = p.id
    GROUP BY
      p.id,
      p.title,
      p.description,
      p.problem_code,
      p.persona,
      p.problem_statement,
      p.solution_idea,
      p.why_selected,
      p.key_technologies,
      p.domain,
      p.is_active,
      p.created_at,
      p.updated_at
    ORDER BY p.problem_code ASC, p.created_at DESC`
  );

  const allSelections = await fetchAllSelections();
  const recentSelections = allSelections.slice(0, 8);

  const [domainSummary] = await pool.query(
    `SELECT
      p.domain,
      COUNT(ps.id) AS teamCount
    FROM problems p
    LEFT JOIN problem_selections ps ON ps.problem_id = p.id
    WHERE p.is_active = 1
    GROUP BY p.domain
    ORDER BY "teamCount" DESC, p.domain ASC`
  );

  const problems = problemsRaw.map(normalizeProblem);
  const activeProblems = problems.filter((problem) => Number(problem.isActive) === 1);
  const fullProblems = activeProblems.filter((problem) => problem.isFull);

  return {
    problems,
    recentSelections,
    allSelections,
    domainSummary,
    domains: Array.from(new Set([...DEFAULT_DOMAINS, ...problems.map((problem) => problem.domain)])).sort(),
    settings: await fetchSettings(),
    stats: {
      totalProblems: problems.length,
      activeProblems: activeProblems.length,
      totalSelections: problems.reduce((sum, problem) => sum + problem.teamCount, 0),
      fullProblems: fullProblems.length,
    },
  };
}

app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));
app.set("trust proxy", 1);

app.use(express.urlencoded({ extended: false }));
app.use(express.static(path.join(__dirname, "public")));

app.use(
  session({
    key: "hackathon.sid",
    secret: process.env.SESSION_SECRET || "replace-this-secret",
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      sameSite: "lax",
      secure: isProduction,
      maxAge: 1000 * 60 * 60 * 8,
    },
  })
);

app.use((req, res, next) => {
  res.locals.flash = consumeFlash(req);
  res.locals.admin = req.session.admin || null;
  res.locals.participant = req.session.participant || null;
  res.locals.currentYear = new Date().getFullYear();
  res.locals.teamLimit = TEAM_LIMIT;
  res.locals.assetVersion = ASSET_VERSION;
  next();
});

app.get("/", async (req, res, next) => {
  try {
    const publicData = await fetchPublicData(req.session.participant?.id);
    res.render("home", {
      title: "Venture Capital Hackathon 2026 — Malla Reddy University",
      partners: PARTNERS,
      ...publicData,
    });
  } catch (error) {
    next(error);
  }
});

app.post("/participant/signup", async (req, res, next) => {
  const teamName = String(req.body.teamName || "").trim();
  const emailAddress = String(req.body.emailAddress || "").trim().toLowerCase();
  const mobileNumber = String(req.body.mobileNumber || "").trim();
  const password = req.body.password || "";
  const confirmPassword = req.body.confirmPassword || "";

  if (!teamName || !emailAddress || !mobileNumber || !password || !confirmPassword) {
    setFlash(req, "error", "Please complete all signup fields.");
    return res.redirect("/");
  }

  if (!validEmailAddress(emailAddress)) {
    setFlash(req, "error", "Please enter a valid email address.");
    return res.redirect("/");
  }

  if (!validMobileNumber(mobileNumber)) {
    setFlash(req, "error", "Please enter a valid 10-digit mobile number.");
    return res.redirect("/");
  }

  if (!validPassword(password)) {
    setFlash(req, "error", "Password must be at least 6 characters long.");
    return res.redirect("/");
  }

  if (password !== confirmPassword) {
    setFlash(req, "error", "Password and confirm password do not match.");
    return res.redirect("/");
  }

  try {
    const passwordHash = await bcrypt.hash(password, 12);
    const [result] = await pool.query(
      `INSERT INTO participants
        (team_name, team_leader_name, email_address, mobile_number, password_hash)
      VALUES (?, ?, ?, ?, ?) RETURNING id`,
      [teamName, emailAddress, emailAddress, mobileNumber, passwordHash]
    );

    req.session.participant = {
      id: result.insertId,
      teamName,
      emailAddress,
      mobileNumber,
    };

    setFlash(req, "success", `Welcome, ${teamName}. Your participant account is ready.`);
    return res.redirect("/");
  } catch (error) {
    if (error.code === "23505") {
      setFlash(req, "error", "This mobile number is already registered. Please log in instead.");
      return res.redirect("/");
    }

    return next(error);
  }
});

app.post("/participant/login", async (req, res, next) => {
  const mobileNumber = String(req.body.mobileNumber || "").trim();
  const password = req.body.password || "";

  if (!mobileNumber || !password) {
    setFlash(req, "error", "Enter your mobile number and password.");
    return res.redirect("/");
  }

  try {
    const [participants] = await pool.query(
      `SELECT
        id,
        team_name AS teamName,
        COALESCE(NULLIF(email_address, ''), team_leader_name) AS emailAddress,
        mobile_number AS mobileNumber,
        password_hash AS passwordHash
      FROM participants
      WHERE mobile_number = ?
      LIMIT 1`,
      [mobileNumber]
    );

    if (participants.length === 0) {
      setFlash(req, "error", "Invalid participant login details.");
      return res.redirect("/");
    }

    const participant = participants[0];
    const passwordMatches = await bcrypt.compare(password, participant.passwordHash);

    if (!passwordMatches) {
      setFlash(req, "error", "Invalid participant login details.");
      return res.redirect("/");
    }

    req.session.participant = {
      id: participant.id,
      teamName: participant.teamName,
      emailAddress: participant.emailAddress,
      mobileNumber: participant.mobileNumber,
    };

    setFlash(req, "success", `Welcome back, ${participant.teamName}.`);
    return res.redirect("/");
  } catch (error) {
    return next(error);
  }
});

app.post("/participant/logout", requireParticipant, (req, res, next) => {
  delete req.session.participant;
  setFlash(req, "success", "Participant account signed out.");
  return req.session.save((error) => {
    if (error) {
      return next(error);
    }

    return res.redirect("/");
  });
});

app.post("/select", requireParticipant, async (req, res, next) => {
  const problemId = String(req.body.problemId || "").trim();
  const participantId = req.session.participant.id;

  if (!problemId) {
    setFlash(req, "error", "Please choose a problem statement first.");
    return res.redirect("/");
  }

  const connection = await pool.getConnection();
  try {
    await connection.beginTransaction();

    const [participantRows] = await connection.query(
      `SELECT
        id,
        team_name AS teamName,
        COALESCE(NULLIF(email_address, ''), team_leader_name) AS emailAddress,
        mobile_number AS mobileNumber
      FROM participants
      WHERE id = ?
      FOR UPDATE`,
      [participantId]
    );

    if (participantRows.length === 0) {
      throw new Error("Your participant account could not be found. Please sign in again.");
    }

    const participant = participantRows[0];
    const [existingSelections] = await connection.query(
      "SELECT id FROM problem_selections WHERE participant_id = ? FOR UPDATE",
      [participantId]
    );

    if (existingSelections.length > 0) {
      throw new Error("You have already selected a problem statement. Editing is disabled after confirmation.");
    }

    const [problemRows] = await connection.query(
      "SELECT id, title, is_active AS isActive FROM problems WHERE id = ? FOR UPDATE",
      [problemId]
    );

    if (problemRows.length === 0 || Number(problemRows[0].isActive) !== 1) {
      throw new Error("This problem statement is no longer available.");
    }

    const [countRows] = await connection.query(
      "SELECT COUNT(*) AS total FROM problem_selections WHERE problem_id = ?",
      [problemId]
    );

    const currentSelections = Number(countRows[0].total || 0);
    if (currentSelections >= TEAM_LIMIT) {
      throw new Error("This problem statement is already full. Please choose another one.");
    }

    await connection.query(
      `INSERT INTO problem_selections
        (participant_id, problem_id, team_name, team_leader_name, email_address, mobile_number)
      VALUES (?, ?, ?, ?, ?, ?)`,
      [
        participant.id,
        problemId,
        participant.teamName.trim(),
        participant.emailAddress.trim(),
        participant.emailAddress.trim(),
        participant.mobileNumber.trim(),
      ]
    );

    await connection.commit();
    setFlash(
      req,
      "success",
      `Problem statement locked for team "${participant.teamName}". It can no longer be edited.`
    );
    return res.redirect("/");
  } catch (error) {
    await connection.rollback();

    if (error.code === "23505") {
      setFlash(req, "error", "This participant account already has a saved selection.");
      return res.redirect("/");
    }

    setFlash(req, "error", error.message || "We could not save that team selection.");
    return res.redirect("/");
  } finally {
    connection.release();
  }
});

app.get("/admin/login", (req, res) => {
  if (req.session.admin) {
    return res.redirect("/admin");
  }

  return res.render("admin-login", {
    title: "Admin Login",
  });
});

app.post("/admin/login", async (req, res, next) => {
  const username = (req.body.username || "").trim();
  const password = req.body.password || "";

  if (!username || !password) {
    setFlash(req, "error", "Enter your admin username and password.");
    return res.redirect("/admin/login");
  }

  try {
    const [admins] = await pool.query(
      "SELECT id, username, password_hash AS passwordHash FROM admins WHERE username = ? LIMIT 1",
      [username]
    );

    if (admins.length === 0) {
      setFlash(req, "error", "Invalid admin credentials.");
      return res.redirect("/admin/login");
    }

    const admin = admins[0];
    const passwordMatches = await bcrypt.compare(password, admin.passwordHash);

    if (!passwordMatches) {
      setFlash(req, "error", "Invalid admin credentials.");
      return res.redirect("/admin/login");
    }

    req.session.admin = {
      id: admin.id,
      username: admin.username,
    };

    setFlash(req, "success", `Welcome back, ${admin.username}.`);
    return res.redirect("/admin");
  } catch (error) {
    return next(error);
  }
});

app.post("/admin/logout", requireAdmin, (req, res, next) => {
  req.session.destroy((error) => {
    if (error) {
      return next(error);
    }

    res.clearCookie("hackathon.sid");
    return res.redirect("/admin/login");
  });
});

app.get("/admin", requireAdmin, async (req, res, next) => {
  try {
    const adminData = await fetchAdminData();
    res.render("admin-dashboard", {
      title: "Admin Dashboard",
      ...adminData,
    });
  } catch (error) {
    next(error);
  }
});

app.get("/admin/selections/export.csv", requireAdmin, async (req, res, next) => {
  try {
    const selections = await fetchAllSelections();
    const headers = [
      "Team Name",
      "Email Address",
      "Mobile Number",
      "Problem Statement",
      "Domain",
      "Selected At",
    ];

    const rows = selections.map((item) => [
      item.teamName,
      item.emailAddress,
      item.mobileNumber,
      item.problemTitle,
      item.problemDomain,
      item.createdAtFormatted,
    ]);

    const csv = [headers, ...rows]
      .map((row) => row.map(escapeCsvValue).join(","))
      .join("\n");

    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    res.setHeader("Content-Disposition", 'attachment; filename="team-selections.csv"');
    res.send(csv);
  } catch (error) {
    next(error);
  }
});

app.get("/admin/selections/export.xlsx", requireAdmin, async (req, res, next) => {
  try {
    const selections = await fetchAllSelections();
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(
      selections.map((item) => ({
        "Team Name": item.teamName,
        "Email Address": item.emailAddress,
        "Mobile Number": item.mobileNumber,
        "Problem Statement": item.problemTitle,
        Domain: item.problemDomain,
        "Selected At": item.createdAtFormatted,
      }))
    );

    XLSX.utils.book_append_sheet(workbook, worksheet, "Selections");
    const buffer = XLSX.write(workbook, { bookType: "xlsx", type: "buffer" });

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", 'attachment; filename="team-selections.xlsx"');
    res.send(buffer);
  } catch (error) {
    next(error);
  }
});

app.post("/admin/problems", requireAdmin, async (req, res, next) => {
  const {
    problemCode,
    title,
    domain,
    description,
    persona,
    problemStatement,
    solutionIdea,
    whySelected,
    keyTechnologies,
  } = req.body;

  if (!title || !domain || !description) {
    setFlash(req, "error", "Title, domain, and description are all required.");
    return res.redirect("/admin");
  }

  try {
    await pool.query(
      `INSERT INTO problems
        (
          problem_code,
          title,
          domain,
          description,
          persona,
          problem_statement,
          solution_idea,
          why_selected,
          key_technologies
        )
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [
        cleanOptionalText(problemCode),
        title.trim(),
        domain.trim(),
        description.trim(),
        cleanOptionalText(persona),
        cleanOptionalText(problemStatement),
        cleanOptionalText(solutionIdea),
        cleanOptionalText(whySelected),
        cleanOptionalText(keyTechnologies),
      ]
    );

    setFlash(req, "success", "Problem statement created successfully.");
    return res.redirect("/admin");
  } catch (error) {
    return next(error);
  }
});

app.post("/admin/problems/:id/update", requireAdmin, async (req, res, next) => {
  const { id } = req.params;
  const {
    problemCode,
    title,
    domain,
    description,
    persona,
    problemStatement,
    solutionIdea,
    whySelected,
    keyTechnologies,
    isActive,
  } = req.body;

  if (!title || !domain || !description) {
    setFlash(req, "error", "Please complete the title, domain, and description fields.");
    return res.redirect("/admin");
  }

  try {
    await pool.query(
      `UPDATE problems
      SET
        problem_code = ?,
        title = ?,
        domain = ?,
        description = ?,
        persona = ?,
        problem_statement = ?,
        solution_idea = ?,
        why_selected = ?,
        key_technologies = ?,
        is_active = ?
      WHERE id = ?`,
      [
        cleanOptionalText(problemCode),
        title.trim(),
        domain.trim(),
        description.trim(),
        cleanOptionalText(persona),
        cleanOptionalText(problemStatement),
        cleanOptionalText(solutionIdea),
        cleanOptionalText(whySelected),
        cleanOptionalText(keyTechnologies),
        isActive === "1" ? 1 : 0,
        id,
      ]
    );

    setFlash(req, "success", "Problem statement updated.");
    return res.redirect("/admin");
  } catch (error) {
    return next(error);
  }
});

app.post("/admin/problems/:id/delete", requireAdmin, async (req, res, next) => {
  try {
    await pool.query("DELETE FROM problems WHERE id = ?", [req.params.id]);
    setFlash(req, "success", "Problem statement deleted.");
    return res.redirect("/admin");
  } catch (error) {
    return next(error);
  }
});

app.post("/admin/selections/:id/delete", requireAdmin, async (req, res, next) => {
  try {
    const [result] = await pool.query("DELETE FROM problem_selections WHERE id = ?", [req.params.id]);

    if (result.affectedRows === 0) {
      setFlash(req, "error", "That team selection could not be found.");
      return res.redirect("/admin");
    }

    setFlash(req, "success", "Team selection deleted.");
    return res.redirect("/admin");
  } catch (error) {
    return next(error);
  }
});

app.post("/admin/selections/:id/update", requireAdmin, async (req, res, next) => {
  const selectionId = String(req.params.id || "").trim();
  const teamName = String(req.body.teamName || "").trim();
  const emailAddress = String(req.body.emailAddress || "").trim().toLowerCase();
  const mobileNumber = String(req.body.mobileNumber || "").trim();
  const problemId = String(req.body.problemId || "").trim();

  if (!teamName || !emailAddress || !mobileNumber || !problemId) {
    setFlash(req, "error", "Please complete all selection edit fields.");
    return res.redirect("/admin");
  }

  if (!validEmailAddress(emailAddress)) {
    setFlash(req, "error", "Please enter a valid email address.");
    return res.redirect("/admin");
  }

  if (!validMobileNumber(mobileNumber)) {
    setFlash(req, "error", "Please enter a valid 10-digit mobile number.");
    return res.redirect("/admin");
  }

  const connection = await pool.getConnection();
  try {
    await connection.beginTransaction();

    const [selectionRows] = await connection.query(
      `SELECT id, participant_id AS participantId
      FROM problem_selections
      WHERE id = ?
      FOR UPDATE`,
      [selectionId]
    );

    if (selectionRows.length === 0) {
      throw new Error("That team selection could not be found.");
    }

    const selection = selectionRows[0];

    const [problemRows] = await connection.query(
      `SELECT id
      FROM problems
      WHERE id = ?
      LIMIT 1`,
      [problemId]
    );

    if (problemRows.length === 0) {
      throw new Error("The selected problem statement could not be found.");
    }

    await connection.query(
      `UPDATE problem_selections
      SET
        problem_id = ?,
        team_name = ?,
        team_leader_name = ?,
        email_address = ?,
        mobile_number = ?
      WHERE id = ?`,
      [problemId, teamName, emailAddress, emailAddress, mobileNumber, selectionId]
    );

    if (selection.participantId) {
      await connection.query(
        `UPDATE participants
        SET
          team_name = ?,
          team_leader_name = ?,
          email_address = ?,
          mobile_number = ?
        WHERE id = ?`,
        [teamName, emailAddress, emailAddress, mobileNumber, selection.participantId]
      );
    }

    await connection.commit();
    setFlash(req, "success", "Admin updated the team selection successfully.");
    return res.redirect("/admin");
  } catch (error) {
    await connection.rollback();

    if (error.code === "23505") {
      setFlash(req, "error", "That email or mobile number is already in use by another record.");
      return res.redirect("/admin");
    }

    setFlash(req, "error", error.message || "Could not update that team selection.");
    return res.redirect("/admin");
  } finally {
    connection.release();
  }
});

app.get("/health", async (req, res) => {
  try {
    await pool.query("SELECT 1");
    res.json({ ok: true });
  } catch (error) {
    res.status(500).json({ ok: false, error: error.message });
  }
});

app.post("/admin/settings/update", requireAdmin, async (req, res, next) => {
  try {
    const { whatsapp_url, linkedin_url, instagram_url } = req.body;

    await pool.query("UPDATE settings SET setting_value = ? WHERE setting_key = 'whatsapp_url'", [whatsapp_url || ""]);
    await pool.query("UPDATE settings SET setting_value = ? WHERE setting_key = 'linkedin_url'", [linkedin_url || ""]);
    await pool.query("UPDATE settings SET setting_value = ? WHERE setting_key = 'instagram_url'", [instagram_url || ""]);

    res.redirect("/admin");
  } catch (error) {
    next(error);
  }
});

app.use((error, req, res, next) => {
  console.error(error);
  res.status(500).render("error", {
    title: "Something went wrong",
    message: error.message || "Unexpected server error.",
  });
});

async function startServer() {
  await initializeDatabase();
  await ensureSettingsTable();
  await ensureProblemBriefColumns();
  await ensureParticipantSelectionSupport();
  await ensureParticipantEmailColumns();
  await ensureAdminAccount();
  await syncDefaultProblemCatalog();

  app.listen(PORT, () => {
    console.log(`Hackathon selection app running on http://localhost:${PORT}`);
  });
}

startServer().catch((error) => {
  console.error("Failed to start server:", error);
  process.exit(1);
});

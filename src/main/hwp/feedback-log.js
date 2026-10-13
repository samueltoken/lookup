const fs = require("node:fs");
const fsp = require("node:fs/promises");
const path = require("node:path");

const PROBLEMS_HEADER = "# Problems\n\n";
const SOLUTIONS_HEADER = "# Solutions\n\n";

function escapeLine(value) {
  return String(value || "").replace(/\r?\n/g, " ").trim();
}

function hasAsarSegment(value) {
  const raw = String(value || "");
  if (!raw) {
    return false;
  }
  const normalized = raw.replace(/\//g, "\\");
  return normalized
    .split(/\\+/)
    .filter(Boolean)
    .some((segment) => /\.asar$/i.test(segment));
}

async function ensureMarkdownLogFile(filePath, headerText) {
  try {
    if (!filePath || typeof filePath !== "string") {
      return false;
    }
    const resolved = path.resolve(filePath);
    if (hasAsarSegment(resolved)) {
      return false;
    }
    const dirPath = path.dirname(resolved);
    await fsp.mkdir(dirPath, { recursive: true });
    if (!fs.existsSync(resolved)) {
      await fsp.writeFile(resolved, headerText, "utf8");
      return true;
    }
    const current = await fsp.readFile(resolved, "utf8").catch(() => "");
    if (!String(current || "").trim()) {
      await fsp.writeFile(resolved, headerText, "utf8");
    }
    return true;
  } catch (_error) {
    return false;
  }
}

function resolveProjectRoot(appRoot) {
  const candidate = path.resolve(appRoot || process.cwd());
  if (hasAsarSegment(candidate)) {
    const parentDir = path.dirname(candidate);
    try {
      if (!hasAsarSegment(parentDir) && fs.existsSync(parentDir) && fs.statSync(parentDir).isDirectory()) {
        return parentDir;
      }
    } catch (_error) {
      // ignore and fallback
    }
    const cwd = path.resolve(process.cwd());
    if (!hasAsarSegment(cwd)) {
      return cwd;
    }
    return path.resolve(process.env.USERPROFILE || "");
  }
  try {
    if (fs.existsSync(candidate) && fs.statSync(candidate).isDirectory()) {
      return candidate;
    }
  } catch (_error) {
    // ignore and fallback
  }
  return process.cwd();
}

function buildLogPaths(appRoot) {
  const root = resolveProjectRoot(appRoot);
  const globalRoot = path.resolve(process.env.USERPROFILE || "", ".codex", "skills", "_global_feedback");
  return {
    projectProblems: path.join(root, "problems.md"),
    projectSolutions: path.join(root, "solutions.md"),
    globalProblems: path.join(globalRoot, "problems.md"),
    globalSolutions: path.join(globalRoot, "solutions.md")
  };
}

async function ensureAndReadLogs(paths) {
  await Promise.allSettled([
    ensureMarkdownLogFile(paths.projectProblems, PROBLEMS_HEADER),
    ensureMarkdownLogFile(paths.projectSolutions, SOLUTIONS_HEADER),
    ensureMarkdownLogFile(paths.globalProblems, PROBLEMS_HEADER),
    ensureMarkdownLogFile(paths.globalSolutions, SOLUTIONS_HEADER)
  ]);

  const [projectProblems, projectSolutions, globalProblems, globalSolutions] = await Promise.all([
    fsp.readFile(paths.projectProblems, "utf8").catch(() => ""),
    fsp.readFile(paths.projectSolutions, "utf8").catch(() => ""),
    fsp.readFile(paths.globalProblems, "utf8").catch(() => ""),
    fsp.readFile(paths.globalSolutions, "utf8").catch(() => "")
  ]);

  return {
    projectProblems,
    projectSolutions,
    globalProblems,
    globalSolutions
  };
}

async function appendLogEntry(filePath, key, sectionText, headerText) {
  const marker = `[feedback-key:${String(key || "").trim()}]`;
  if (!String(marker).trim() || marker === "[feedback-key:]") {
    return;
  }
  const resolved = path.resolve(filePath);
  const ready = await ensureMarkdownLogFile(resolved, headerText || PROBLEMS_HEADER);
  if (!ready) {
    return;
  }
  try {
    const current = await fsp.readFile(resolved, "utf8").catch(() => "");
    if (String(current || "").includes(marker)) {
      return;
    }
    const payload = `${sectionText}\n\n${marker}\n`;
    await fsp.appendFile(resolved, payload, "utf8");
  } catch (_error) {
    // ignore log write failures
  }
}

function buildProblemSection(nowIso, title, requestSummary, issueSummary, rootCause, improvementPlan) {
  return [
    `## [${nowIso}] ${escapeLine(title) || "문제 기록"}`,
    `- 요청 요약: ${escapeLine(requestSummary) || "(없음)"}`,
    `- 불편 내용: ${escapeLine(issueSummary) || "(없음)"}`,
    `- 원인: ${escapeLine(rootCause) || "(미분석)"}`,
    `- 개선 방향: ${escapeLine(improvementPlan) || "(미정)"}`
  ].join("\n");
}

function buildSolutionSection(nowIso, title, changesMade, qualityGain, satisfactionPoint) {
  return [
    `## [${nowIso}] ${escapeLine(title) || "개선 기록"}`,
    `- 무엇을 고쳤는지: ${escapeLine(changesMade) || "(없음)"}`,
    `- 어떻게 좋아졌는지: ${escapeLine(qualityGain) || "(없음)"}`,
    `- 만족 포인트: ${escapeLine(satisfactionPoint) || "(없음)"}`
  ].join("\n");
}

function createFeedbackLogger(options = {}) {
  const paths = buildLogPaths(options.appRoot);
  let warmupDone = false;

  const warmup = async () => {
    const logs = await ensureAndReadLogs(paths);
    warmupDone = true;
    return logs;
  };

  const ensureWarmup = async () => {
    if (!warmupDone) {
      await warmup();
    }
  };

  const recordProblem = async ({ key, title, requestSummary, issueSummary, rootCause, improvementPlan }) => {
    await ensureWarmup();
    const nowIso = new Date().toISOString().slice(0, 10);
    const section = buildProblemSection(nowIso, title, requestSummary, issueSummary, rootCause, improvementPlan);
    await appendLogEntry(paths.projectProblems, key, section, PROBLEMS_HEADER);
    await appendLogEntry(paths.globalProblems, key, section, PROBLEMS_HEADER);
  };

  const recordSolution = async ({ key, title, changesMade, qualityGain, satisfactionPoint }) => {
    await ensureWarmup();
    const nowIso = new Date().toISOString().slice(0, 10);
    const section = buildSolutionSection(nowIso, title, changesMade, qualityGain, satisfactionPoint);
    await ensureMarkdownLogFile(paths.projectSolutions, SOLUTIONS_HEADER);
    await ensureMarkdownLogFile(paths.globalSolutions, SOLUTIONS_HEADER);
    await appendLogEntry(paths.projectSolutions, key, section, SOLUTIONS_HEADER);
    await appendLogEntry(paths.globalSolutions, key, section, SOLUTIONS_HEADER);
  };

  return {
    paths,
    warmup,
    recordProblem,
    recordSolution
  };
}

module.exports = {
  createFeedbackLogger
};

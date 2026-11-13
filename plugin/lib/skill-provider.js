/**
 * 把 `plugin/skills/*\/SKILL.md` 暴露给 DSH 的技能目录。
 *
 * DSH 不会自动扫描插件包的 skills 目录：必须在 apply() 里通过
 * `ctx.inject(['skills'], ctx => ctx.skills.registerProvider(...))` 注册一个
 * provider。provider 的契约（与 dsh-ui-ux-pro-max 保持一致）：
 *
 *   list(options) -> candidate[]   // 目录里出现的候选
 *   get(candidate, options) -> { ...candidate, content }  // 取正文
 *
 * candidate 关键字段：name / description / whenToUse / invocation /
 * provider / source / resourceBase / path / locator / rank / metadata。
 *
 * 这里只用 Node 内置模块，不引入第三方依赖，宿主老版本也能安全降级。
 */

import { promises as fs } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const SKILLS_DIR = path.resolve(HERE, '..', 'skills');

const PROVIDER_NAME = 'dsh-word-pipeline';
const INVOCATION = { modelInvocable: true, userInvocable: true };

/** frontmatter 只解析标量字段：name / description / whenToUse。 */
export function splitFrontmatter(text) {
  const source = String(text ?? '').replace(/^\uFEFF/, '');
  const match = /^---\r?\n([\s\S]*?)\r?\n---\r?\n?/.exec(source);
  if (!match) return { data: {}, body: source };
  const data = {};
  for (const rawLine of match[1].split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith('#')) continue;
    const sep = line.indexOf(':');
    if (sep <= 0) continue;
    const key = line.slice(0, sep).trim();
    let value = line.slice(sep + 1).trim();
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    if (key) data[key] = value;
  }
  return { data, body: source.slice(match[0].length) };
}

let cachedRank;
/** 排序权重：宿主导出 BUNDLED_SKILL_RANK 时用它的值，取不到就用中性值。 */
async function resolveRank() {
  if (cachedRank !== undefined) return cachedRank;
  try {
    const mod = await import('@deepseek-ai/dsh-skill');
    cachedRank = mod.BUNDLED_SKILL_RANK ?? 50;
  } catch {
    cachedRank = 50;
  }
  return cachedRank;
}

function candidateFrom(file, dirName, data, rank) {
  return {
    name: data.name || dirName,
    description: data.description || '',
    whenToUse: data.whenToUse || '',
    invocation: INVOCATION,
    provider: PROVIDER_NAME,
    source: 'bundled',
    resourceBase: { kind: 'directory', path: path.dirname(file) },
    path: file,
    ...(rank === undefined ? {} : { rank }),
    locator: file,
    metadata: data,
  };
}

async function readSkill(file, dirName) {
  const text = await fs.readFile(file, 'utf8');
  const { data, body } = splitFrontmatter(text);
  return { file, dirName, data, body };
}

export const skillProvider = {
  name: PROVIDER_NAME,

  async list(options = {}) {
    if (options.signal?.aborted) return [];
    let entries;
    try {
      entries = await fs.readdir(SKILLS_DIR, { withFileTypes: true });
    } catch {
      return [];
    }
    const rank = await resolveRank();
    const candidates = [];
    for (const entry of entries) {
      if (!entry.isDirectory()) continue;
      const dirName = entry.name;
      const file = path.join(SKILLS_DIR, dirName, 'SKILL.md');
      try {
        const { data } = await readSkill(file, dirName);
        candidates.push(candidateFrom(file, dirName, data, rank));
      } catch {
        // 目录里没有 SKILL.md 或读不了：跳过，不影响其它技能
      }
    }
    return candidates;
  },

  async get(candidate, options = {}) {
    if (options.signal?.aborted) return undefined;
    const file = candidate?.locator || candidate?.path;
    if (!file) return undefined;
    const dirName = path.basename(path.dirname(file));
    const { data, body } = await readSkill(file, dirName);
    const rank = await resolveRank();
    return {
      ...candidateFrom(file, dirName, data, rank),
      name: candidate?.name || data.name || dirName,
      description: data.description || candidate?.description || '',
      whenToUse: data.whenToUse || candidate?.whenToUse || '',
      metadata: data,
      content: body,
    };
  },
};

export const SKILL_NAMES = ['docx-retypeset', 'word-pipeline'];
export const SKILLS_ROOT = SKILLS_DIR;

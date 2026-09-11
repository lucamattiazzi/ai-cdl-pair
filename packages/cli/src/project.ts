import { access, mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { createJiti } from "jiti";
import { z } from "zod";
import { aiCdlConfigSchema, type ResolvedAiCdlConfig } from "./config.js";

/** Loaded config plus resolved project paths. */
export interface LoadedProject {
  readonly root: string;
  readonly configPath: string;
  readonly config: ResolvedAiCdlConfig;
  readonly appId: string;
}

/** Load, validate, and resolve the generated app identity for an AI-CDL project. */
export async function loadProject(cwd = process.cwd(), persistId = true): Promise<LoadedProject> {
  const configPath = resolve(cwd, "ai-cdl.config.ts");
  try {
    await access(configPath);
  } catch {
    throw new Error("AI_CDL_CONFIG_NOT_FOUND: Create ai-cdl.config.ts or run `ai-cdl init`.");
  }
  const jiti = createJiti(configPath, { interopDefault: true });
  const imported = await jiti.import(configPath, { default: true });
  const parsed = aiCdlConfigSchema.safeParse(imported);
  if (!parsed.success) {
    const issues = z.prettifyError(parsed.error);
    throw new Error(`AI_CDL_CONFIG_INVALID\n${issues}\nFix the listed paths in ai-cdl.config.ts.`);
  }
  let appId = parsed.data.app.id;
  if (appId === "generate") {
    const idPath = resolve(cwd, ".ai-cdl/app-id");
    try {
      appId = (await readFile(idPath, "utf8")).trim();
      z.string().uuid().parse(appId);
    } catch {
      if (!persistId)
        throw new Error(
          "AI_CDL_APP_ID_MISSING: Generate the persistent app ID before running a drift check.",
        );
      appId = crypto.randomUUID();
      await mkdir(dirname(idPath), { recursive: true });
      await writeFile(idPath, `${appId}\n`, { flag: "wx" }).catch(async (error: unknown) => {
        if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error;
        appId = (await readFile(idPath, "utf8")).trim();
      });
    }
  }
  return { root: cwd, configPath, config: parsed.data, appId };
}

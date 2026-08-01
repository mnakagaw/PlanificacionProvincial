import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import { Client } from "basic-ftp";
import dotenv from "dotenv";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const projectDir = path.resolve(scriptDir, "..");
const envPath = process.env.FTP_ENV_FILE
  ? path.resolve(process.env.FTP_ENV_FILE)
  : path.resolve(projectDir, "..", ".env");

dotenv.config({ path: envPath, quiet: true });

const required = ["FTP_HOST", "FTP_USER", "FTP_PASS", "FTP_REMOTE_ROOT"];
const missing = required.filter((key) => !process.env[key]);
if (missing.length) {
  throw new Error(`Missing FTP settings: ${missing.join(", ")}`);
}

const uploadDir = path.resolve(projectDir, "dist");
await fs.access(path.join(uploadDir, "index.html"));

const configuredRoot = process.env.FTP_REMOTE_ROOT
  .replaceAll("\\", "/")
  .replace(/\/+$/, "");
const ddptRoot = configuredRoot.slice(0, configuredRoot.lastIndexOf("/"));
const targetPath =
  process.env.FTP_TARGET_PATH?.replaceAll("\\", "/").replace(/\/+$/, "") ||
  `${ddptRoot}/PlanificacionProvincial`;

if (!/\/DDPT\/PlanificacionProvincial$/i.test(targetPath)) {
  throw new Error("Refusing FTP deploy outside DDPT/PlanificacionProvincial.");
}

const client = new Client(30_000);

try {
  await client.access({
    host: process.env.FTP_HOST,
    user: process.env.FTP_USER,
    password: process.env.FTP_PASS,
    secure: process.env.FTP_SECURE === "true",
  });
  await client.ensureDir(targetPath);
  await client.clearWorkingDir();
  await client.uploadFromDir(uploadDir);
  const publishedFiles = await client.list();
  console.log(
    `Published ${publishedFiles.length} top-level items to DDPT/PlanificacionProvincial.`,
  );
} finally {
  client.close();
}

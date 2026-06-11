const fs = require("fs/promises");
const path = require("path");
const { gzipSync, gunzipSync } = require("zlib");
const { S3Client, GetObjectCommand, PutObjectCommand } = require("@aws-sdk/client-s3");

const R2_OBJECT_KEY = process.env.R2_OBJECT_KEY || "fuel-ets/dashboard-state.json";
const LOCAL_STATE_PATH = process.env.STATE_FILE_PATH || path.join(__dirname, ".runtime", "dashboard-state.json");
const STATE_CACHE_MS = Number(process.env.STATE_CACHE_MS || 30_000);
let cachedDocument;
let cachedAt = 0;
let loadPromise = null;

function r2Configured() {
  return Boolean(
    process.env.R2_ACCOUNT_ID &&
      process.env.R2_ACCESS_KEY_ID &&
      process.env.R2_SECRET_ACCESS_KEY &&
      process.env.R2_BUCKET_NAME
  );
}

function storageMode() {
  return r2Configured() ? "r2" : "file";
}

function createR2Client() {
  return new S3Client({
    region: "auto",
    endpoint: `https://${process.env.R2_ACCOUNT_ID}.r2.cloudflarestorage.com`,
    credentials: {
      accessKeyId: process.env.R2_ACCESS_KEY_ID,
      secretAccessKey: process.env.R2_SECRET_ACCESS_KEY,
    },
  });
}

async function bodyToBuffer(body) {
  if (!body) return Buffer.alloc(0);
  if (typeof body.transformToByteArray === "function") {
    return Buffer.from(await body.transformToByteArray());
  }
  if (typeof body.transformToString === "function") {
    return Buffer.from(await body.transformToString());
  }

  const chunks = [];
  for await (const chunk of body) {
    chunks.push(Buffer.from(chunk));
  }
  return Buffer.concat(chunks);
}

function normalizeDocument(value) {
  if (!value || typeof value !== "object" || !value.state) {
    return null;
  }
  return {
    schemaVersion: Number(value.schemaVersion) || 1,
    revision: String(value.revision || ""),
    updatedAt: String(value.updatedAt || ""),
    updatedBy: String(value.updatedBy || ""),
    state: value.state,
  };
}

async function readFromR2() {
  const client = createR2Client();
  try {
    const response = await client.send(
      new GetObjectCommand({
        Bucket: process.env.R2_BUCKET_NAME,
        Key: R2_OBJECT_KEY,
      })
    );
    const body = await bodyToBuffer(response.Body);
    const text = response.ContentEncoding === "gzip" || (body[0] === 0x1f && body[1] === 0x8b)
      ? gunzipSync(body).toString("utf8")
      : body.toString("utf8");
    return normalizeDocument(JSON.parse(text));
  } catch (error) {
    if (error?.name === "NoSuchKey" || error?.$metadata?.httpStatusCode === 404) {
      return null;
    }
    throw error;
  }
}

async function writeToR2(document) {
  const client = createR2Client();
  await client.send(
    new PutObjectCommand({
      Bucket: process.env.R2_BUCKET_NAME,
      Key: R2_OBJECT_KEY,
      Body: gzipSync(Buffer.from(JSON.stringify(document))),
      ContentType: "application/json",
      ContentEncoding: "gzip",
      CacheControl: "no-store",
    })
  );
}

async function readFromFile() {
  try {
    return normalizeDocument(JSON.parse(await fs.readFile(LOCAL_STATE_PATH, "utf8")));
  } catch (error) {
    if (error.code === "ENOENT") return null;
    throw error;
  }
}

async function writeToFile(document) {
  await fs.mkdir(path.dirname(LOCAL_STATE_PATH), { recursive: true });
  const temporaryPath = `${LOCAL_STATE_PATH}.${process.pid}.tmp`;
  await fs.writeFile(temporaryPath, JSON.stringify(document), "utf8");
  await fs.rename(temporaryPath, LOCAL_STATE_PATH);
}

async function loadStateDocument() {
  if (cachedDocument !== undefined && Date.now() - cachedAt < STATE_CACHE_MS) {
    return cachedDocument;
  }
  if (loadPromise) return loadPromise;

  loadPromise = (r2Configured() ? readFromR2() : readFromFile())
    .then((document) => {
      cachedDocument = document;
      cachedAt = Date.now();
      return document;
    })
    .finally(() => {
      loadPromise = null;
    });
  return loadPromise;
}

async function saveStateDocument(state, updatedBy = "") {
  const now = new Date().toISOString();
  const document = {
    schemaVersion: 1,
    revision: `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`,
    updatedAt: now,
    updatedBy: String(updatedBy || ""),
    state,
  };

  if (r2Configured()) {
    await writeToR2(document);
  } else {
    await writeToFile(document);
  }
  cachedDocument = document;
  cachedAt = Date.now();
  return document;
}

function getStorageStatus() {
  return {
    mode: storageMode(),
    durable: r2Configured(),
    configured: r2Configured(),
    objectKey: r2Configured() ? R2_OBJECT_KEY : path.basename(LOCAL_STATE_PATH),
  };
}

module.exports = {
  getStorageStatus,
  loadStateDocument,
  saveStateDocument,
};

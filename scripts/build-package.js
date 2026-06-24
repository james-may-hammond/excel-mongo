/**
 * File: build-package.js
 * Description: Build script to package the frontend or backend assets.
 * Dependencies: fs, path, child_process
 */
const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const dist = path.join(root, "dist");
const sourceFe = path.join(root, "fe");
const localUrl = "https://localhost:8000";

function requiredUrl(name) {
    let value = process.env[name];
    
    // Auto-detect Vercel deployment URLs
    if (name === "ADDIN_PUBLIC_URL" && !value) {
        if (process.env.VERCEL_PROJECT_PRODUCTION_URL) {
            value = "https://" + process.env.VERCEL_PROJECT_PRODUCTION_URL;
        } else if (process.env.VERCEL_URL) {
            value = "https://" + process.env.VERCEL_URL;
        }
    }

    if (!value) {
        throw new Error(`${name} is required. Example: ${name}=https://excel-mongo.example.com npm run package`);
    }

    // Auto-prepend https:// if the user just pasted the domain (like "excel-mongo.up.railway.app")
    if (!value.startsWith("http://") && !value.startsWith("https://")) {
        value = "https://" + value;
    }

    let parsed;
    try {
        parsed = new URL(value);
    } catch {
        throw new Error(`${name} must be a valid URL. Received: ${value}`);
    }

    if (parsed.protocol !== "https:") {
        throw new Error(`${name} must use https for Office add-in production deployment.`);
    }

    return value.replace(/\/+$/, "");
}

function optionalUrl(name, fallback) {
    const value = process.env[name];
    if (!value) return fallback;
    return requiredUrl(name);
}

function copyDir(source, target) {
    fs.mkdirSync(target, { recursive: true });
    for (const entry of fs.readdirSync(source, { withFileTypes: true })) {
        const from = path.join(source, entry.name);
        const to = path.join(target, entry.name);
        if (entry.isDirectory()) {
            copyDir(from, to);
        } else {
            fs.copyFileSync(from, to);
        }
    }
}

function manifestVersion(version) {
    const parts = version.split(".").map((part) => Number.parseInt(part, 10));
    while (parts.length < 4) parts.push(0);
    return parts.slice(0, 4).map((part) => Number.isFinite(part) ? part : 0).join(".");
}

const publicUrl = requiredUrl("ADDIN_PUBLIC_URL");
const apiBase = optionalUrl("ADDIN_API_BASE", publicUrl);
const supportUrl = optionalUrl("ADDIN_SUPPORT_URL", publicUrl);
const publicOrigin = new URL(publicUrl).origin;
const apiOrigin = new URL(apiBase).origin;
const packageJson = require(path.join(root, "package.json"));

fs.rmSync(dist, { recursive: true, force: true });
copyDir(sourceFe, path.join(dist, "fe"));
fs.rmSync(path.join(dist, "fe", "manifest.xml"), { force: true });

const configPath = path.join(dist, "fe", "config.js");
fs.writeFileSync(
    configPath,
    `window.EXCEL_MONGO_CONFIG = ${JSON.stringify({ apiBase }, null, 4)};\n`,
    "utf8"
);

const manifestSource = fs.readFileSync(path.join(sourceFe, "manifest.xml"), "utf8");
let manifest = manifestSource
    .replaceAll(localUrl, publicUrl)
    .replace(/<SupportUrl DefaultValue="[^"]*" \/>/, `<SupportUrl DefaultValue="${supportUrl}" />`)
    .replace(/<Version>[^<]+<\/Version>/, `<Version>${manifestVersion(packageJson.version)}</Version>`);

if (apiOrigin !== publicOrigin && !manifest.includes(`<AppDomain>${apiOrigin}</AppDomain>`)) {
    manifest = manifest.replace(
        "</AppDomains>",
        `  <AppDomain>${apiOrigin}</AppDomain>\n  </AppDomains>`
    );
}

fs.writeFileSync(path.join(dist, "manifest.xml"), manifest, "utf8");

console.log(`Packaged Office add-in into ${path.relative(root, dist)}/`);
console.log(`Manifest: ${path.relative(root, path.join(dist, "manifest.xml"))}`);
console.log(`Static assets: ${path.relative(root, path.join(dist, "fe"))}/`);
console.log(`Taskpane URL: ${publicUrl}/fe/taskpane.html`);
console.log(`API base: ${apiBase}`);

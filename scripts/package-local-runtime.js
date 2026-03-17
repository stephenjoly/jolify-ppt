const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");

const repoRoot = path.resolve(__dirname, "..");
const distDir = path.join(repoRoot, "dist");
const stagingRoot = path.join(repoRoot, ".local-runtime-build");
const bundleDir = path.join(stagingRoot, "bundle");
const webDir = path.join(bundleDir, "web");
const bundleArchive = path.join(distDir, "jolify-local-bundle.tar.gz");
const sourceManifestPath = path.join(repoRoot, "manifest.xml");
const localServerPath = path.join(repoRoot, "scripts", "local_server.py");
const LOCAL_BASE_URL = "https://127.0.0.1:38443/";
const LOCAL_APP_DOMAIN = "https://127.0.0.1:38443";

function ensureExists(targetPath) {
  if (!fs.existsSync(targetPath)) {
    throw new Error(`Expected build artifact not found: ${targetPath}`);
  }
}

function copyRecursive(source, destination) {
  fs.cpSync(source, destination, { recursive: true });
}

function buildLocalManifest() {
  const source = fs.readFileSync(sourceManifestPath, "utf8");
  let localManifest = source.replace(/https:\/\/localhost:3300\//g, LOCAL_BASE_URL);

  if (!localManifest.includes(`<AppDomain>${LOCAL_APP_DOMAIN}</AppDomain>`)) {
    localManifest = localManifest.replace(
      /<\/AppDomains>/,
      `    <AppDomain>${LOCAL_APP_DOMAIN}</AppDomain>\n  </AppDomains>`,
    );
  }

  return localManifest;
}

function main() {
  ensureExists(distDir);
  ensureExists(localServerPath);

  fs.rmSync(stagingRoot, { recursive: true, force: true });
  fs.rmSync(bundleArchive, { force: true });

  fs.mkdirSync(webDir, { recursive: true });

  [
    "assets",
    "dialogs",
    "commands.html",
    "commands.js",
    "commands.js.map",
    "taskpane.html",
    "taskpane.js",
    "taskpane.js.map",
    "polyfill.js",
    "polyfill.js.map",
  ].forEach((entry) => {
    const sourcePath = path.join(distDir, entry);
    ensureExists(sourcePath);
    const destinationPath = path.join(webDir, entry);
    copyRecursive(sourcePath, destinationPath);
  });

  const localManifest = buildLocalManifest();
  fs.writeFileSync(path.join(distDir, "manifest.local.xml"), localManifest);
  fs.writeFileSync(path.join(bundleDir, "manifest.xml"), localManifest);
  fs.copyFileSync(localServerPath, path.join(bundleDir, "local-server.py"));
  fs.writeFileSync(
    path.join(bundleDir, "version.json"),
    JSON.stringify({ builtAt: new Date().toISOString() }, null, 2),
  );

  execFileSync("tar", ["-czf", bundleArchive, "-C", bundleDir, "."], { stdio: "inherit" });

  fs.rmSync(stagingRoot, { recursive: true, force: true });
  console.log(`Created local runtime bundle: ${bundleArchive}`);
}

main();

// Bundles the task pane and the auth dialog with esbuild and copies static files to dist/.
import { build, context } from "esbuild";
import { cpSync, mkdirSync, rmSync } from "node:fs";

const watch = process.argv.includes("--watch");
const outdir = "dist";

rmSync(outdir, { recursive: true, force: true });
mkdirSync(outdir, { recursive: true });

const options = {
  entryPoints: {
    taskpane: "src/taskpane.js",
    dialog: "src/dialog.js",
  },
  bundle: true,
  format: "iife",
  target: ["es2020"],
  minify: !watch,
  sourcemap: watch ? "inline" : false,
  outdir,
  logLevel: "info",
};

// Static files: HTML, CSS, icons and the config sample
for (const file of ["taskpane.html", "taskpane.css", "dialog.html"]) {
  cpSync(`src/${file}`, `${outdir}/${file}`);
}
cpSync("public", outdir, { recursive: true });

if (watch) {
  const ctx = await context(options);
  await ctx.watch();
  console.log("Watching for changes …");
} else {
  await build(options);
}

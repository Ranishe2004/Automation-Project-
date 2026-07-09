import { readFileSync, writeFileSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));

const part1 = readFileSync(join(__dirname, "__part1.txt"), "utf-8");
const part2 = readFileSync(join(__dirname, "__part2.txt"), "utf-8");

const content = part1 + "\n" + part2;

const outPath = join(__dirname, "App.jsx");
writeFileSync(outPath, content, "utf-8");

console.log("Written: " + outPath);
console.log("Size: " + content.length + " bytes");

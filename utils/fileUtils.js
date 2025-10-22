import fs from "fs";
import path from "path";

export function getSavePath(type, format = "docx") {
  const folder = path.resolve(`./generated/${type}`);
  if (!fs.existsSync(folder)) fs.mkdirSync(folder, { recursive: true });

  const files = fs.readdirSync(folder).filter(f => f.endsWith(`.${format}`));
  const counter = files.length + 1;
  const filename = `${type}_Lease_${counter}.${format}`;
  return path.join(folder, filename);
}

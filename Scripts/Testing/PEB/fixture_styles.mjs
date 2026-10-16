import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";

let definitions;
export function fixtureStyle(name) {
  if (!definitions) {
    const python = process.env.PEB_PYTHON;
    if (!python) throw new Error("Set PEB_PYTHON to the bundled Python executable.");
    const bridge = fileURLToPath(new URL("./xlsx_expectations_bridge.py", import.meta.url));
    definitions = JSON.parse(execFileSync(python, ["-X", "utf8", bridge, "fixture-styles"], { encoding: "utf8" }));
  }
  const resolve = (key, path = []) => {
    if (!definitions[key] || path.includes(key)) throw new Error(`Missing or cyclic fixture style: ${key}`);
    const { basedOn, ...attributes } = definitions[key];
    return { ...(basedOn ? resolve(basedOn, [...path, key]) : {}), ...attributes };
  };
  return resolve(name);
}

export function applyFixtureStyle(range, name) {
  const style = fixtureStyle(name);
  const format = {};
  for (const [attribute, property] of Object.entries({ backColor: "fill", horizontal: "horizontalAlignment", vertical: "verticalAlignment" })) {
    if (style[attribute] !== undefined) format[property] = style[attribute];
  }
  const font = {};
  for (const [attribute, property] of Object.entries({ fontName: "name", textColor: "color", fontSize: "size", fontBold: "bold" })) {
    if (style[attribute] !== undefined) font[property] = attribute === "fontSize" ? Number(style[attribute]) : attribute === "fontBold" ? style[attribute] === "true" : style[attribute];
  }
  if (Object.keys(font).length) format.font = font;
  if (style.overflow !== undefined) format.wrapText = style.overflow === "wrap";
  for (const attribute of ["columnWidth", "rowHeight"]) {
    if (style[attribute] !== undefined) format[attribute] = Number(style[attribute]);
  }
  range.format = format;
  if (style.numberFormat !== undefined) range.setNumberFormat(style.numberFormat);
}
"""Inspect the UTF-16 recovery journal without modifying Excel or its workbooks."""

import argparse
import json
from pathlib import Path
from xml.etree import ElementTree as ET


def inspect(path):
    prepared = {}
    committed = []
    invalid_lines = []
    for number, line in enumerate(path.read_text(encoding="utf-16", errors="replace").splitlines(), 1):
        if not line.strip():
            continue
        try:
            root = ET.fromstring(line)
            if root.tag != "Recovery" or root.get("version") != "1":
                raise ValueError("Unknown journal format")
            token = root.attrib["operation"]
            state = root.attrib["state"]
            if state == "PREPARE":
                if token in prepared:
                    raise ValueError("Duplicate operation identifier")
                prepared[token] = {
                    "operation": token,
                    "workbook": root.attrib["workbook"],
                    "time": root.attrib["time"],
                    "user": root.attrib["user"],
                    "rows": [
                        {
                            "role": row.attrib["role"],
                            "sheet": row.attrib["sheet"],
                            "table": row.attrib["table"],
                            "fields": {
                                field.attrib["name"]: {
                                    "type": int(field.attrib["type"]),
                                    "value": field.text or "",
                                }
                                for field in row.findall("Field")
                            },
                        }
                        for row in root.findall("Row")
                    ],
                }
            elif state == "COMMIT":
                committed.append(prepared.pop(token))
            else:
                raise ValueError("Unknown operation state")
        except (ET.ParseError, KeyError, ValueError):
            invalid_lines.append(number)
    return {"committed": committed, "pending": list(prepared.values()), "invalid_lines": invalid_lines}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("journal", type=Path)
    args = parser.parse_args()
    print(json.dumps(inspect(args.journal), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
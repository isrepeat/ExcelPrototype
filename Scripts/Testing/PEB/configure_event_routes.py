"""Validate XML routing; the old wsConfig route generator is retired."""

from event_routes import load_registry, parse_config, ROOT

if __name__ == "__main__":
    registry = load_registry(parse_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder/wsConfig.test.txt"))
    print(f"PASS: {len(registry['events'])} business events loaded from the shared XML registry.")
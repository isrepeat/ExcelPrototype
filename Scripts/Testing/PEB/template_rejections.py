"""Validate optional export rejection rules shared by Word and Movement."""

import re

NS = {"p": "urn:excelprototype:profiles"}


def validate_rejections(template, settings=None, known=None):
    for rule in template.findall("p:reject", NS):
        if set(rule.attrib) not in ({"field", "equals", "messageKey"}, {"field", "matches", "messageKey"}):
            raise ValueError("Rejection requires field, messageKey and exactly one of equals or matches")
        if any(not value.strip() for value in rule.attrib.values()) or len(rule) or rule.text:
            raise ValueError("Rejection must be an empty element with nonempty attributes")
        if known is not None and rule.get("field") not in known:
            raise ValueError("Unknown rejection field: " + rule.get("field"))
        if settings is not None and not settings.get(rule.get("messageKey"), "").strip():
            raise ValueError("Missing rejection message: " + rule.get("messageKey"))
        if "matches" in rule.attrib:
            try:
                re.compile(rule.get("matches"))
            except re.error as error:
                raise ValueError("Invalid rejection pattern: " + str(error)) from error
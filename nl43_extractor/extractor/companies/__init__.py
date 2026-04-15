from .bajaj_allianz import parse_bajaj_allianz

PARSER_REGISTRY: dict = {
    "parse_bajaj_allianz": parse_bajaj_allianz,
}

from parsers import bel_parser, unknown_parser, gem_parser

# Add new parsers here as you build them
PARSER_MAP = {
    "BEL":     bel_parser,
    "GEM":     gem_parser,
    "UNKNOWN": unknown_parser,
    # "HAL":  hal_parser,
    # "DRDO": drdo_parser,
}

def route(text: str, company_code: str, pdf_bytes: bytes = None) -> list[dict]:
    """
    Selects the correct parser based on company_code and runs it.
    Falls back to unknown_parser if company_code not in map.
    For GEM, passes pdf_bytes to the parser.
    """
    parser = PARSER_MAP.get(company_code, unknown_parser)
    if company_code == "GEM":
        return parser.parse(text, pdf_bytes=pdf_bytes)
    return parser.parse(text)

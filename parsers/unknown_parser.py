def parse(text: str) -> list[dict]:
    """
    Fallback parser for unknown companies.
    Returns a single row with NA for all fields and a note in Remark.
    """
    from parsers.bel_parser import COLUMNS
    row = {col: "NA" for col in COLUMNS}
    row["Remark"] = "Unknown company — manual review needed"
    row["Status"] = "Pending Review"
    return [row]

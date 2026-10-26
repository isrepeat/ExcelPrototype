"""Validate the configured fixed medical-report layout without Office."""

from event_routes import ROOT, parse_config


def main():
    expected = (
        "RowNumber|Unit|Rank|PersonName|Dates|DocumentName|DocumentDate|"
        "DocumentNumber|CombatWound|Hospital|Diagnosis|TreatmentStatus|"
        "DischargeDocumentName|DischargeDocumentDate|DischargeDocumentNumber|"
        "Phone|PersonId|BirthDate|Certificates"
    ).split("|")
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        config = parse_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename)
        prefix = "validation.MedicalReports."
        columns = config[prefix + "source.columns"].split("|")
        assert columns == expected
        assert len(columns) == len(set(columns)) == 19
        assert config[prefix + "source.range"] == "A3:S20000"
        assert set(config[prefix + "fields"].split("|")) <= set(columns)
    source = (ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder/ex_PEB_MedicalReports.vba").read_text(encoding="utf-8-sig")
    read_file = source.split("Private Function private_ReadFile(", 1)[1].split("Private Function private_CheckPeople(", 1)[0]
    assert 'private_HeaderIndex' not in read_file
    assert 'indexes("RowNumber")' in read_file
    assert 'If VBA.IsNull(value) Or VBA.IsEmpty(value) Then Exit For' in read_file
    assert 'For row = 1 To dataRows' in read_file
    for stage in ('Columns', 'Bounds', 'DateFilter', 'BuildResults', 'Filter'):
        assert 'MedicalValidation.File.' + stage in read_file
    first_date = source.split('Private Function private_FirstDateMatches(', 1)[1].split('Private Function private_HeaderIndex(', 1)[0]
    assert 'dateRx' not in first_date
    assert 'VBA.Mid$(text, index, 10) <> orderDateText' in first_date
    assert 'character >= 48 And character <= 57' in first_date
    assert 'character <> 160' in first_date
    print("PASS: PROD/TEST fixed A:S layout, numbered-block boundary and separate timing stages. VBA was not executed.")


if __name__ == "__main__":
    main()
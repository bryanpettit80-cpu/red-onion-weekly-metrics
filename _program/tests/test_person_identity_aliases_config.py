from __future__ import annotations

import pytest

from red_onion_config import ConfigError, validate_config_payload


def test_person_aliases_are_explicit_location_specific_and_trimmed() -> None:
    config = validate_config_payload({
        "person_identity_aliases": {
            " RC Richmond ": {" Taylor G. ": " Taylor Guest "},
        },
    })
    assert config["person_identity_aliases"] == {
        "RC Richmond": {"Taylor G.": "Taylor Guest"},
    }
    assert config["public_name_aliases"]["Bar Server"] == "Bar"


@pytest.mark.parametrize(
    "aliases,message",
    [
        ({"Unknown Store": {"Old Name": "Current Name"}}, "not configured"),
        ({"RC Richmond": []}, "must be a JSON object"),
        ({"RC Richmond": {"Old Name": "Old Name"}}, "different raw names"),
        ({"RC Richmond": {"Old Name": "Middle Name", "Middle Name": "New Name"}}, "chains and cycles"),
        ({"RC Richmond": {"Old Name": "New Name", "New Name": "Old Name"}}, "chains and cycles"),
        ({"RC Richmond": {"Old Name": "New Name", " Old Name ": "Another Name"}}, "unique after trimming"),
        ({"RC Richmond": {"5050 Server Server2": "Taylor Guest"}}, "shared POS"),
        ({"RC Richmond": {"12 Server1 Server2": "Taylor Guest"}}, "shared POS"),
        ({"RC Richmond": {"Taylor Guest": "5050 Bar"}}, "shared POS"),
        ({"RC Richmond": {"Taylor Guest": "Server 2"}}, "shared POS"),
        ({"RC Richmond": {"123": "Taylor Guest"}}, "shared POS"),
    ],
)
def test_ambiguous_or_nonperson_alias_configuration_is_rejected(aliases, message: str) -> None:
    with pytest.raises(ConfigError, match=message):
        validate_config_payload({"person_identity_aliases": aliases})


def test_unconfigured_alias_mapping_does_not_guess_name_equivalence() -> None:
    assert validate_config_payload({})["person_identity_aliases"] == {}

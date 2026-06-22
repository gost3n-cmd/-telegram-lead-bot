"""Тесты для email_validator (раздел 11 архитектуры — pytest)."""

import pytest

from bot.validation.email_validator import validate_email, ValidationResult


class TestValidationResult:
    def test_create_valid(self):
        r = ValidationResult(is_valid=True, email="a@b.com")
        assert r.is_valid is True
        assert r.email == "a@b.com"
        assert r.error is None

    def test_create_invalid(self):
        r = ValidationResult(is_valid=False, error="bad")
        assert r.is_valid is False
        assert r.email is None
        assert r.error == "bad"


class TestValidateEmailNormalization:
    """Раздел 7: нормализация (нижний регистр, обрезка пробелов)."""

    def test_trim_whitespace(self):
        r = validate_email("  user@example.com  ")
        assert r.is_valid
        assert r.email == "user@example.com"

    def test_lowercase(self):
        r = validate_email("USER@EXAMPLE.COM")
        assert r.is_valid
        assert r.email == "user@example.com"

    def test_mixed_case_and_spaces(self):
        r = validate_email("  John.DOE@Example.COM  ")
        assert r.is_valid
        assert r.email == "john.doe@example.com"

    def test_plus_tag(self):
        r = validate_email("user+tag@example.com")
        assert r.is_valid
        assert r.email == "user+tag@example.com"

    def test_subdomain(self):
        r = validate_email("user@sub.example.co.uk")
        assert r.is_valid
        assert r.email == "user@sub.example.co.uk"


class TestValidateEmailInvalid:
    """Раздел 13.2: невалидный синтаксис."""

    def test_empty_string(self):
        assert validate_email("").is_valid is False

    def test_none(self):
        assert validate_email(None).is_valid is False

    def test_no_at_sign(self):
        assert validate_email("plainaddress").is_valid is False

    def test_empty_local(self):
        assert validate_email("@example.com").is_valid is False

    def test_empty_domain(self):
        assert validate_email("user@").is_valid is False

    def test_no_dot_in_domain(self):
        assert validate_email("user@example").is_valid is False

    def test_spaces_inside(self):
        assert validate_email("a b@example.com").is_valid is False

    def test_single_char_tld_rejected(self):
        # DNS не имеет односимвольных TLD
        assert validate_email("user@example.c").is_valid is False

    def test_too_long_email(self):
        long_local = "a" * 300
        assert validate_email(f"{long_local}@example.com").is_valid is False

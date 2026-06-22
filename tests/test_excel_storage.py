"""Тесты для ExcelStorage (раздел 11 архитектуры — pytest, pytest-asyncio).

Покрывает StorageInterface контракт:
- save_lead, email_exists, count_leads, list_leads, health_check
- Дубликаты, невалидный ввод, создание файла при отсутствии.
"""

import os
import tempfile

import pytest

from bot.models.lead import Lead
from bot.services.lead_service import LeadService, ServiceErrorType
from bot.storage.excel_storage import ExcelStorage


@pytest.fixture
def temp_xlsx():
    """Временный xlsx-файл, удаляется после теста."""
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "test.xlsx")
        yield path


@pytest.fixture
def storage(temp_xlsx):
    """ExcelStorage с временным файлом."""
    return ExcelStorage(file_path=temp_xlsx, sheet_name="Leads")


@pytest.fixture
def service(storage):
    """LeadService с тестовым storage."""
    return LeadService(storage=storage)


class TestExcelStorageContract:
    """Контракт StorageInterface (раздел 10)."""

    @pytest.mark.asyncio
    async def test_save_lead_creates_file_when_missing(self, storage, temp_xlsx):
        """Раздел 4.5: бот создаёт Leads при первом лиде, если файла нет."""
        assert not os.path.exists(temp_xlsx)
        lead = Lead(telegram_id=1, email="a@example.com")
        result = await storage.save_lead(lead)
        assert result.success
        assert os.path.exists(temp_xlsx)

    @pytest.mark.asyncio
    async def test_save_lead_returns_success(self, storage):
        lead = Lead(telegram_id=42, email="test@example.com", telegram_username="tester")
        result = await storage.save_lead(lead)
        assert result.success
        assert result.error is None

    @pytest.mark.asyncio
    async def test_email_exists_false_initially(self, storage):
        assert await storage.email_exists("nobody@example.com") is False

    @pytest.mark.asyncio
    async def test_email_exists_true_after_save(self, storage):
        lead = Lead(telegram_id=1, email="found@example.com")
        await storage.save_lead(lead)
        assert await storage.email_exists("found@example.com") is True

    @pytest.mark.asyncio
    async def test_email_exists_case_insensitive(self, storage):
        lead = Lead(telegram_id=1, email="mixed@Example.COM")
        await storage.save_lead(lead)
        assert await storage.email_exists("MIXED@example.com") is True

    @pytest.mark.asyncio
    async def test_count_leads_empty(self, storage):
        assert await storage.count_leads() == 0

    @pytest.mark.asyncio
    async def test_count_leads_after_saves(self, storage):
        await storage.save_lead(Lead(telegram_id=1, email="a@example.com"))
        await storage.save_lead(Lead(telegram_id=2, email="b@example.com"))
        await storage.save_lead(Lead(telegram_id=3, email="c@example.com"))
        assert await storage.count_leads() == 3

    @pytest.mark.asyncio
    async def test_list_leads_pagination(self, storage):
        for i in range(5):
            await storage.save_lead(Lead(telegram_id=i, email=f"u{i}@example.com"))
        page1 = await storage.list_leads(limit=2, offset=0)
        page2 = await storage.list_leads(limit=2, offset=2)
        assert len(page1) == 2
        assert len(page2) == 2

    @pytest.mark.asyncio
    async def test_health_check_true_when_available(self, storage):
        assert await storage.health_check() is True

    @pytest.mark.asyncio
    async def test_health_check_false_when_corrupt(self, temp_xlsx):
        # Создать повреждённый файл
        with open(temp_xlsx, "w") as f:
            f.write("not a valid xlsx")
        storage = ExcelStorage(file_path=temp_xlsx, sheet_name="Leads")
        assert await storage.health_check() is False


class TestLeadServiceIntegration:
    """Интеграционные тесты LeadService + ExcelStorage (Application + Infrastructure)."""

    @pytest.mark.asyncio
    async def test_process_email_success(self, service):
        result = await service.process_email(
            raw_email="alice@example.com",
            telegram_id=111,
            telegram_username="alice",
        )
        assert result.success
        assert result.lead is not None
        assert result.lead.email == "alice@example.com"

    @pytest.mark.asyncio
    async def test_process_email_normalizes(self, service):
        result = await service.process_email(
            raw_email="  BOB@EXAMPLE.COM  ",
            telegram_id=222,
        )
        assert result.success
        assert result.lead.email == "bob@example.com"

    @pytest.mark.asyncio
    async def test_process_email_invalid(self, service):
        result = await service.process_email(
            raw_email="not-an-email",
            telegram_id=333,
        )
        assert not result.success
        assert result.error_type == ServiceErrorType.VALIDATION_FAILED

    @pytest.mark.asyncio
    async def test_process_email_duplicate(self, service):
        await service.process_email("carol@example.com", telegram_id=1)
        result = await service.process_email("carol@example.com", telegram_id=2)
        assert not result.success
        assert result.error_type == ServiceErrorType.DUPLICATE

    @pytest.mark.asyncio
    async def test_concurrent_saves_no_lost_update(self, storage):
        """Раздел 14: конкурентные записи не должны теряться (asyncio.Lock)."""
        import asyncio

        service = LeadService(storage=storage)
        emails = [f"u{i}@example.com" for i in range(10)]

        await asyncio.gather(
            *[service.process_email(e, telegram_id=i) for i, e in enumerate(emails)]
        )

        count = await storage.count_leads()
        assert count == 10, f"Lost update! Expected 10, got {count}"

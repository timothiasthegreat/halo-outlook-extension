"""Tests for HaloClient with mocked HTTP responses and TokenManager."""

import tempfile
from unittest.mock import AsyncMock, MagicMock

import pytest
from respx import MockRouter

from watcher.crypto import generate_key
from watcher.halo_client import HaloClient
from watcher.state import StateStore
from watcher.token_manager import TokenManager

API_URL = "https://test.halopsa.com/api"


@pytest.fixture
def fernet_key():
    return generate_key()


@pytest.fixture
async def get_state():
    db_path = tempfile.mktemp(suffix=".db")
    async with StateStore(db_path) as s:
        yield s


@pytest.fixture
async def token_manager(fernet_key):
    """Create a TokenManager that returns a fixed access token without HTTP calls."""
    async with StateStore(":memory:") as state:
        tm = TokenManager(
            {"instance_url": "https://test.halopsa.com", "client_id": "test-client"},
            state,
            fernet_key,
        )
        # Register a test user and seed the cache
        await tm.register_user("test@example.com", "dummy-refresh-token")
        # Manually inject a cached token to bypass HTTP
        tm._cache["test@example.com"] = ("mock-access-token", float("inf"))
        yield tm


@pytest.fixture
def mock_api():
    """Mock Halo API endpoints (not the auth endpoint)."""
    with MockRouter(assert_all_called=False) as respx_mock:
        yield respx_mock


@pytest.mark.asyncio
class TestHaloClientTickets:
    async def test_create_ticket(self, token_manager, mock_api):
        mock_api.post("https://test.halopsa.com/api/Tickets").respond(
            json=[{"id": 9999, "summary": "Test ticket"}]
        )

        async with HaloClient(token_manager, "test@example.com", API_URL) as client:
            ticket = await client.create_ticket(
                summary="Test ticket",
                details_html="<p>body</p>",
                tickettype_id=1,
            )

        assert ticket["id"] == 9999

    async def test_create_ticket_with_custom_field(self, token_manager, mock_api):
        mock_api.post("https://test.halopsa.com/api/Tickets").respond(
            json=[{"id": 9999}]
        )

        async with HaloClient(token_manager, "test@example.com", API_URL) as client:
            ticket = await client.create_ticket(
                summary="Test",
                details_html="<p>test</p>",
                customfield_value="conv-abc-123",
            )

        assert ticket["id"] == 9999

    async def test_get_ticket(self, token_manager, mock_api):
        mock_api.get("https://test.halopsa.com/api/Tickets/9999").respond(
            json={"id": 9999, "summary": "Test", "status_id": 1}
        )

        async with HaloClient(token_manager, "test@example.com", API_URL) as client:
            ticket = await client.get_ticket(9999)

        assert ticket["id"] == 9999
        assert ticket["status_id"] == 1

    async def test_update_ticket_custom_field(self, token_manager, mock_api):
        mock_api.post("https://test.halopsa.com/api/Tickets").respond(
            json=[{"id": 9999}]
        )

        async with HaloClient(token_manager, "test@example.com", API_URL) as client:
            result = await client.update_ticket_custom_field(9999, "new-conv-id")

        assert result is not None

    async def test_search_tickets(self, token_manager, mock_api):
        mock_api.get("https://test.halopsa.com/api/Tickets").respond(
            json={"record_count": 1, "tickets": [{"id": 9999, "summary": "Found it"}]}
        )

        async with HaloClient(token_manager, "test@example.com", API_URL) as client:
            results = await client.search_tickets("9999")

        assert len(results) == 1
        assert results[0]["id"] == 9999


@pytest.mark.asyncio
class TestHaloClientActions:
    async def test_create_action(self, token_manager, mock_api):
        mock_api.post("https://test.halopsa.com/api/Actions").respond(
            json=[{"id": 42, "ticket_id": 9999, "outcome_id": 16}]
        )

        async with HaloClient(token_manager, "test@example.com", API_URL) as client:
            action = await client.create_action(
                ticket_id=9999,
                outcome_id=16,
                note="Test note",
                note_html="<p>Test note</p>",
                email_message_id="test-msg@example.com",
            )

        assert action["id"] == 42
        assert action["ticket_id"] == 9999

    async def test_create_action_internal_note(self, token_manager, mock_api):
        mock_api.post("https://test.halopsa.com/api/Actions").respond(
            json=[{"id": 43, "ticket_id": 9999, "outcome_id": 7}]
        )

        async with HaloClient(token_manager, "test@example.com", API_URL) as client:
            action = await client.create_action(
                ticket_id=9999,
                outcome_id=7,
                note="Internal note",
                note_html="<p>Internal</p>",
                hiddenfromuser=True,
            )

        assert action["outcome_id"] == 7

    async def test_list_actions(self, token_manager, mock_api):
        mock_api.get("https://test.halopsa.com/api/Actions").respond(
            json=[
                {"id": 1, "ticket_id": 9999, "outcome_id": 0},
                {"id": 2, "ticket_id": 9999, "outcome_id": 16},
            ]
        )

        async with HaloClient(token_manager, "test@example.com", API_URL) as client:
            actions = await client.list_actions(9999)

        assert len(actions) == 2


@pytest.mark.asyncio
class TestHaloClientRetry:
    async def test_retries_on_500(self, token_manager):
        import httpx as httpx_mod
        with MockRouter(assert_all_called=False) as respx_mock:
            call_count = 0

            def side_effect(request):
                nonlocal call_count
                call_count += 1
                if call_count <= 2:
                    return httpx_mod.Response(500)
                return httpx_mod.Response(200, json={"id": 1})

            respx_mock.get("https://test.halopsa.com/api/Tickets/1").mock(
                side_effect=side_effect
            )

            async with HaloClient(token_manager, "test@example.com", API_URL) as client:
                result = await client.get_ticket(1)

            assert result["id"] == 1
            assert call_count == 3
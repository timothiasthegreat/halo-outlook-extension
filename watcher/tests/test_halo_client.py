"""Tests for HaloClient with mocked HTTP responses."""

import pytest
from respx import MockRouter

from watcher.config import HaloConfig
from watcher.halo_client import HaloClient


@pytest.fixture
def halo_config():
    return HaloConfig.model_validate(
        {
            "instance_url": "https://test.halopsa.com",
            "client_id": "test-client",
            "client_secret": "test-secret",
        }
    )


@pytest.fixture
def mock_halo():
    """Mock Halo API with token endpoint and basic ticket responses."""
    with MockRouter(assert_all_called=False) as respx_mock:
        # Token endpoint
        respx_mock.post("https://test.halopsa.com/auth/token").respond(
            json={
                "access_token": "mock-access-token",
                "refresh_token": "mock-refresh-token",
                "expires_in": 3600,
                "token_type": "Bearer",
            }
        )
        yield respx_mock


@pytest.mark.asyncio
class TestHaloClientTickets:
    async def test_create_ticket(self, halo_config, mock_halo):
        mock_halo.post("https://test.halopsa.com/api/Tickets").respond(
            json=[{"id": 9999, "summary": "Test ticket"}]
        )

        async with HaloClient(halo_config) as client:
            ticket = await client.create_ticket(
                summary="Test ticket",
                details_html="<p>body</p>",
                tickettype_id=1,
            )

        assert ticket["id"] == 9999

    async def test_create_ticket_with_custom_field(self, halo_config, mock_halo):
        mock_halo.post("https://test.halopsa.com/api/Tickets").respond(
            json=[{"id": 9999}]
        )

        async with HaloClient(halo_config) as client:
            ticket = await client.create_ticket(
                summary="Test",
                details_html="<p>test</p>",
                customfield_value="conv-abc-123",
            )

        assert ticket["id"] == 9999

    async def test_get_ticket(self, halo_config, mock_halo):
        mock_halo.get("https://test.halopsa.com/api/Tickets/9999").respond(
            json={"id": 9999, "summary": "Test", "status_id": 1}
        )

        async with HaloClient(halo_config) as client:
            ticket = await client.get_ticket(9999)

        assert ticket["id"] == 9999
        assert ticket["status_id"] == 1

    async def test_update_ticket_custom_field(self, halo_config, mock_halo):
        mock_halo.post("https://test.halopsa.com/api/Tickets").respond(
            json=[{"id": 9999}]
        )

        async with HaloClient(halo_config) as client:
            result = await client.update_ticket_custom_field(9999, "new-conv-id")

        assert result is not None

    async def test_search_tickets(self, halo_config, mock_halo):
        mock_halo.get("https://test.halopsa.com/api/Tickets").respond(
            json={"record_count": 1, "tickets": [{"id": 9999, "summary": "Found it"}]}
        )

        async with HaloClient(halo_config) as client:
            results = await client.search_tickets("9999")

        assert len(results) == 1
        assert results[0]["id"] == 9999


@pytest.mark.asyncio
class TestHaloClientActions:
    async def test_create_action(self, halo_config, mock_halo):
        mock_halo.post("https://test.halopsa.com/api/Actions").respond(
            json=[{"id": 42, "ticket_id": 9999, "outcome_id": 16}]
        )

        async with HaloClient(halo_config) as client:
            action = await client.create_action(
                ticket_id=9999,
                outcome_id=16,
                note="Test note",
                note_html="<p>Test note</p>",
                email_message_id="test-msg@example.com",
            )

        assert action["id"] == 42
        assert action["ticket_id"] == 9999

    async def test_create_action_internal_note(self, halo_config, mock_halo):
        mock_halo.post("https://test.halopsa.com/api/Actions").respond(
            json=[{"id": 43, "ticket_id": 9999, "outcome_id": 7}]
        )

        async with HaloClient(halo_config) as client:
            action = await client.create_action(
                ticket_id=9999,
                outcome_id=7,
                note="Internal note",
                note_html="<p>Internal</p>",
                hiddenfromuser=True,
            )

        assert action["outcome_id"] == 7

    async def test_list_actions(self, halo_config, mock_halo):
        mock_halo.get("https://test.halopsa.com/api/Actions").respond(
            json=[
                {"id": 1, "ticket_id": 9999, "outcome_id": 0},
                {"id": 2, "ticket_id": 9999, "outcome_id": 16},
            ]
        )

        async with HaloClient(halo_config) as client:
            actions = await client.list_actions(9999)

        assert len(actions) == 2


@pytest.mark.asyncio
class TestHaloClientRetry:
    async def test_retries_on_500(self, halo_config):
        import httpx as httpx_mod
        with MockRouter(assert_all_called=False) as respx_mock:
            respx_mock.post("https://test.halopsa.com/auth/token").respond(
                json={"access_token": "t", "expires_in": 3600}
            )
            # First two requests fail with 500, third succeeds
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

            async with HaloClient(halo_config) as client:
                result = await client.get_ticket(1)

            assert result["id"] == 1
            assert call_count == 3
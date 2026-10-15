"""Tests for GraphClient with mocked HTTP responses."""

import pytest
from respx import MockRouter

from watcher.config import GraphConfig
from watcher.graph_client import GraphClient


@pytest.fixture
def graph_config():
    return GraphConfig.model_validate(
        {
            "tenant_id": "test-tenant",
            "client_id": "test-client",
            "client_secret": "test-secret",
            "user_email": "test@example.com",
        }
    )


@pytest.fixture
def mock_graph():
    """Mock Graph API with token endpoint."""
    with MockRouter(assert_all_called=False) as respx_mock:
        respx_mock.post(
            "https://login.microsoftonline.com/test-tenant/oauth2/v2.0/token"
        ).respond(
            json={
                "access_token": "mock-graph-token",
                "expires_in": 3600,
            }
        )
        yield respx_mock


@pytest.mark.asyncio
class TestGraphClientMessages:
    async def test_get_messages_by_conversation(self, graph_config, mock_graph):
        mock_graph.get(
            "https://graph.microsoft.com/v1.0/users/test@example.com/messages"
        ).respond(
            json={
                "value": [
                    {
                        "subject": "Re: Test",
                        "from": {"emailAddress": {"address": "customer@example.com"}},
                        "sentDateTime": "2026-08-10T12:00:00Z",
                        "internetMessageId": "<msg1@example.com>",
                        "conversationId": "conv-123",
                    },
                    {
                        "subject": "Re: Test",
                        "from": {"emailAddress": {"address": "test@example.com"}},
                        "sentDateTime": "2026-08-10T12:05:00Z",
                        "internetMessageId": "<msg2@example.com>",
                        "conversationId": "conv-123",
                    },
                ]
            }
        )

        async with GraphClient(graph_config) as client:
            messages = await client.get_messages_by_conversation("conv-123")

        assert len(messages) == 2
        # Verify both directions are present
        senders = [m["from"]["emailAddress"]["address"] for m in messages]
        assert "customer@example.com" in senders
        assert "test@example.com" in senders

    async def test_get_messages_by_conversation_empty(self, graph_config, mock_graph):
        mock_graph.get(
            "https://graph.microsoft.com/v1.0/users/test@example.com/messages"
        ).respond(json={"value": []})

        async with GraphClient(graph_config) as client:
            messages = await client.get_messages_by_conversation("conv-empty")

        assert len(messages) == 0

    async def test_get_messages_since(self, graph_config, mock_graph):
        mock_graph.get(
            "https://graph.microsoft.com/v1.0/users/test@example.com/messages"
        ).respond(
            json={
                "value": [
                    {
                        "subject": "Recent",
                        "from": {"emailAddress": {"address": "someone@example.com"}},
                        "sentDateTime": "2026-08-10T14:00:00Z",
                        "conversationId": "conv-recent",
                    }
                ]
            }
        )

        async with GraphClient(graph_config) as client:
            messages = await client.get_messages_since("2026-08-10T00:00:00Z")

        assert len(messages) == 1
        assert messages[0]["subject"] == "Recent"

    async def test_get_single_message(self, graph_config, mock_graph):
        mock_graph.get(
            "https://graph.microsoft.com/v1.0/users/test@example.com/messages/msg-id-123"
        ).respond(
            json={
                "id": "msg-id-123",
                "subject": "Single Message",
                "body": {"content": "<p>Body</p>"},
            }
        )

        async with GraphClient(graph_config) as client:
            msg = await client.get_message("msg-id-123")

        assert msg["id"] == "msg-id-123"
        assert msg["subject"] == "Single Message"


@pytest.mark.asyncio
class TestGraphClientPagination:
    async def test_single_page(self, graph_config, mock_graph):
        mock_graph.get(
            "https://graph.microsoft.com/v1.0/users/test@example.com/messages"
        ).respond(
            json={
                "value": [{"subject": f"Msg {i}"} for i in range(3)],
                # No @odata.nextLink → single page
            }
        )

        async with GraphClient(graph_config) as client:
            messages = await client.get_messages_by_conversation("conv-test")

        assert len(messages) == 3

    async def test_paginates_multiple_pages(self, graph_config, mock_graph):
        """Pages 2+ follow @odata.nextLink without doubling the /v1.0 prefix."""
        import httpx as httpx_mod

        messages_url = "https://graph.microsoft.com/v1.0/users/test@example.com/messages"
        next_link = (
            messages_url
            + "?%24filter=conversationId+eq+%27conv-test%27&%24top=50&%24skip=50"
        )
        pages = {
            None: {
                "value": [{"subject": f"Msg {i}"} for i in range(50)],
                "@odata.nextLink": next_link,
            },
            "50": {"value": [{"subject": f"Msg {i}"} for i in range(50, 60)]},
        }
        requested = []

        def respond(request):
            requested.append(request.url)
            return httpx_mod.Response(200, json=pages[request.url.params.get("$skip")])

        # The old bug: page 2 went to /v1.0/v1.0/... and was dropped
        doubled = mock_graph.get(
            url__startswith="https://graph.microsoft.com/v1.0/v1.0/"
        ).respond(404)
        route = mock_graph.get(messages_url).mock(side_effect=respond)

        async with GraphClient(graph_config) as client:
            messages = await client.get_messages_by_conversation("conv-test")

        assert not doubled.called
        assert route.call_count == 2
        assert requested[1].path == "/v1.0/users/test@example.com/messages"
        assert requested[1].params["$skip"] == "50"
        assert len(messages) == 60
        assert messages[-1]["subject"] == "Msg 59"

    async def test_ignores_next_link_to_other_host(self, graph_config, mock_graph):
        """A nextLink off graph.microsoft.com/v1.0 isn't followed (the token stays home)."""
        mock_graph.get(
            "https://graph.microsoft.com/v1.0/users/test@example.com/messages"
        ).respond(
            json={
                "value": [{"subject": "Msg 0"}],
                "@odata.nextLink": "https://evil.example.com/v1.0/users/x/messages?$skip=1",
            }
        )
        evil = mock_graph.get(url__startswith="https://evil.example.com/").respond(200)

        async with GraphClient(graph_config) as client:
            messages = await client.get_messages_by_conversation("conv-test")

        assert len(messages) == 1
        assert not evil.called


@pytest.mark.asyncio
class TestGraphClientAttachments:
    async def test_attachments_use_given_mailbox(self, graph_config, mock_graph):
        base = "https://graph.microsoft.com/v1.0/users/agent@example.com/messages/m1"
        listing = mock_graph.get(f"{base}/attachments").respond(
            json={"value": [{"id": "a1", "name": "x.pdf"}]}
        )
        content = mock_graph.get(f"{base}/attachments/a1/$value").respond(content=b"PDF")

        async with GraphClient(graph_config) as client:
            atts = await client.get_message_attachments("m1", user_email="agent@example.com")
            data = await client.get_attachment_content("m1", "a1", user_email="agent@example.com")

        assert listing.called and content.called
        assert atts == [{"id": "a1", "name": "x.pdf"}]
        assert data == b"PDF"

    async def test_attachments_default_to_configured_mailbox(self, graph_config, mock_graph):
        route = mock_graph.get(
            "https://graph.microsoft.com/v1.0/users/test@example.com/messages/m1/attachments"
        ).respond(json={"value": []})

        async with GraphClient(graph_config) as client:
            await client.get_message_attachments("m1")

        assert route.called


@pytest.mark.asyncio
class TestGraphClientRetry:
    async def test_retries_on_500(self, graph_config):
        import httpx as httpx_mod
        with MockRouter(assert_all_called=False) as respx_mock:
            respx_mock.post(
                "https://login.microsoftonline.com/test-tenant/oauth2/v2.0/token"
            ).respond(json={"access_token": "t", "expires_in": 3600})

            call_count = 0

            def side_effect(request):
                nonlocal call_count
                call_count += 1
                if call_count <= 2:
                    return httpx_mod.Response(500)
                return httpx_mod.Response(200, json={"value": [{"id": "ok"}]})

            respx_mock.get(
                "https://graph.microsoft.com/v1.0/users/test@example.com/messages"
            ).mock(side_effect=side_effect)

            async with GraphClient(graph_config) as client:
                result = await client.get_messages_by_conversation("conv-test")

            assert len(result) == 1
            assert result[0]["id"] == "ok"
            assert call_count == 3
"""Configuration loading and validation for the Halo Outlook Watcher."""

from __future__ import annotations

from pathlib import Path

import yaml
from pydantic import BaseModel, Field, model_validator


class HaloConfig(BaseModel):
    """HaloPSA connection configuration."""

    instance_url: str = Field(description="HaloPSA instance URL, e.g. https://your-instance.halopsa.com")
    client_id: str = Field(min_length=1, description="OAuth2 client ID")
    client_secret: str = Field(default="", description="OAuth2 client secret (not needed for PKCE)")
    actions: HaloActionsConfig = Field(default_factory=lambda: HaloActionsConfig())
    custom_field_conv_id: int = Field(default=285, description="Custom field ID for conversationId")
    default_ticket_type_id: int = Field(default=1, description="Default ticket type for new tickets")

    @property
    def auth_url(self) -> str:
        """Authorization endpoint URL."""
        return f"{self.instance_url.rstrip('/')}/auth"

    @property
    def api_url(self) -> str:
        """API base URL."""
        return f"{self.instance_url.rstrip('/')}/api"

    @property
    def token_url(self) -> str:
        """Token endpoint URL."""
        return f"{self.auth_url}/token"


class HaloActionsConfig(BaseModel):
    """Ticket action outcome IDs — instance-specific."""

    email_received: int = Field(default=0, description="Inbound email from customer")
    email_sent: int = Field(default=16, description="Outbound email to customer")
    internal_note: int = Field(default=7, description="Internal journal note")


class GraphConfig(BaseModel):
    """Microsoft Graph API configuration."""

    tenant_id: str = Field(min_length=1, description="Azure AD tenant ID")
    client_id: str = Field(min_length=1, description="App registration client ID")
    client_secret: str = Field(min_length=1, description="Client secret")
    user_email: str = Field(default="", description="Default mailbox to watch (per-user queries override this)")


class WatcherConfig(BaseModel):
    """Watcher behavior configuration."""

    poll_interval_seconds: int = Field(default=90, ge=30, le=3600)
    stale_conversation_days: int = Field(default=14, ge=1, le=365)
    log_level: str = Field(default="INFO", pattern=r"^(DEBUG|INFO|WARNING|ERROR)$")
    state_db_path: str = Field(default="state.db")


class Config(BaseModel):
    """Root configuration for the Halo Outlook Watcher."""

    halo: HaloConfig
    graph: GraphConfig
    watcher: WatcherConfig = Field(default_factory=WatcherConfig)

    @model_validator(mode="after")
    def validate_urls(self) -> "Config":
        """Ensure URLs are well-formed."""
        if not self.halo.instance_url.startswith("https://"):
            raise ValueError("halo.instance_url must start with https://")
        return self


def load_config(path: str | Path = "config.yaml") -> Config:
    """Load and validate configuration from a YAML file.

    Args:
        path: Path to config.yaml file.

    Returns:
        Validated Config instance.

    Raises:
        FileNotFoundError: If the config file doesn't exist.
        pydantic.ValidationError: If the config fails validation.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Config file not found: {path}")

    with open(path) as f:
        raw = yaml.safe_load(f)

    return Config.model_validate(raw)
"""Tests for config loading and validation."""

import tempfile
from pathlib import Path

import pytest
import yaml

from watcher.config import Config, GraphConfig, HaloConfig, WatcherConfig, load_config


VALID_CONFIG = {
    "halo": {
        "instance_url": "https://your-instance.halopsa.com",
        "client_id": "test-client-id",
        "client_secret": "test-client-secret",
    },
    "graph": {
        "tenant_id": "test-tenant",
        "client_id": "test-graph-client",
        "client_secret": "test-graph-secret",
        "user_email": "test@example.com",
    },
}


def write_temp_config(data: dict) -> Path:
    """Write a temp config YAML file and return its path."""
    tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".yaml", delete=False)
    yaml.dump(data, tmp)
    tmp.close()
    return Path(tmp.name)


class TestHaloConfig:
    def test_minimal_valid(self):
        config = HaloConfig.model_validate(
            {
                "instance_url": "https://test.halopsa.com",
                "client_id": "abc",
                "client_secret": "xyz",
            }
        )
        assert config.instance_url == "https://test.halopsa.com"
        assert config.default_ticket_type_id == 1  # default

    def test_auth_url_derivation(self):
        config = HaloConfig.model_validate(
            {
                "instance_url": "https://test.halopsa.com/",
                "client_id": "abc",
                "client_secret": "xyz",
            }
        )
        assert config.auth_url == "https://test.halopsa.com/auth"
        assert config.api_url == "https://test.halopsa.com/api"
        assert config.token_url == "https://test.halopsa.com/auth/token"

    def test_missing_client_id_rejected(self):
        with pytest.raises(ValueError):
            HaloConfig.model_validate(
                {
                    "instance_url": "https://test.halopsa.com",
                    "client_id": "",
                    "client_secret": "xyz",
                }
            )

    def test_missing_https_rejected_at_root(self):
        """non-https URL validation only runs on the root Config model."""
        import copy
        from pydantic import ValidationError
        data = copy.deepcopy(VALID_CONFIG)
        data["halo"]["instance_url"] = "http://no-tls.com"
        path = write_temp_config(data)
        try:
            with pytest.raises(ValidationError):
                load_config(path)
        finally:
            path.unlink()


class TestGraphConfig:
    def test_valid(self):
        config = GraphConfig.model_validate(
            {
                "tenant_id": "t",
                "client_id": "c",
                "client_secret": "s",
                "user_email": "u@e.com",
            }
        )
        assert config.user_email == "u@e.com"

    def test_empty_user_email_accepted(self):
            """user_email is optional — per-user mailboxes are registered via add-in."""
            config = GraphConfig.model_validate(
                {
                    "tenant_id": "t",
                    "client_id": "c",
                    "client_secret": "s",
                    "user_email": "",
                }
            )
            assert config.user_email == ""


class TestWatcherConfig:
    def test_defaults(self):
        config = WatcherConfig()
        assert config.poll_interval_seconds == 90
        assert config.stale_conversation_days == 14
        assert config.log_level == "INFO"

    def test_custom_values(self):
        config = WatcherConfig.model_validate(
            {"poll_interval_seconds": 120, "stale_conversation_days": 30, "log_level": "DEBUG"}
        )
        assert config.poll_interval_seconds == 120
        assert config.log_level == "DEBUG"


class TestLoadConfig:
    def test_valid_config_loads(self):
        import copy
        data = copy.deepcopy(VALID_CONFIG)
        path = write_temp_config(data)
        try:
            config = load_config(path)
            assert isinstance(config.halo, HaloConfig)
            assert isinstance(config.graph, GraphConfig)
            assert config.halo.instance_url == "https://your-instance.halopsa.com"
        finally:
            path.unlink()

    def test_missing_file_raises(self):
        with pytest.raises(FileNotFoundError):
            load_config("/nonexistent/config.yaml")

    def test_invalid_config_raises_validation_error(self):
        data = dict(VALID_CONFIG)
        data["halo"]["instance_url"] = "http://no-tls.com"  # no https
        path = write_temp_config(data)
        try:
            with pytest.raises(ValueError):
                load_config(path)
        finally:
            path.unlink()


class TestHaloExclusionsConfig:

    @staticmethod
    def _fresh_config():
        """Return a fresh copy of a valid config dict — VALID_CONFIG may be
        mutated by earlier test_invalid_config_raises_validation_error."""
        import copy
        data = copy.deepcopy(VALID_CONFIG)
        data["halo"]["instance_url"] = "https://your-instance.halopsa.com"
        return data

    def test_defaults_empty(self):
        from watcher.config import HaloExclusionsConfig
        cfg = HaloExclusionsConfig()
        assert cfg.ticket_type_ids_create == []
        assert cfg.ticket_type_ids_search == []
        assert cfg.status_ids_search == []

    def test_custom_values(self):
        from watcher.config import HaloExclusionsConfig
        cfg = HaloExclusionsConfig.model_validate({
            "ticket_type_ids_create": [1, 2],
            "ticket_type_ids_search": [3],
            "status_ids_search": [4, 5, 6],
        })
        assert cfg.ticket_type_ids_create == [1, 2]
        assert cfg.status_ids_search == [4, 5, 6]

    def test_yaml_round_trip(self):
        """Verify exclusions load from a full config YAML."""
        data = self._fresh_config()
        data["halo"]["exclusions"] = {"ticket_type_ids_create": [99]}
        path = write_temp_config(data)
        try:
            config = load_config(path)
            assert config.halo.exclusions.ticket_type_ids_create == [99]
            assert config.halo.exclusions.ticket_type_ids_search == []
        finally:
            path.unlink()

    def test_bare_exclusions_key(self):
        """Bare `exclusions:` (YAML null) should not crash — coerced to defaults."""
        data = self._fresh_config()
        data["halo"]["exclusions"] = None
        path = write_temp_config(data)
        try:
            config = load_config(path)
            assert config.halo.exclusions.ticket_type_ids_create == []
            assert config.halo.exclusions.ticket_type_ids_search == []
        finally:
            path.unlink()

    def test_missing_exclusions_key(self):
        """Entire `exclusions` block missing should work with defaults."""
        data = self._fresh_config()
        path = write_temp_config(data)
        try:
            config = load_config(path)
            assert config.halo.exclusions.ticket_type_ids_create == []
            assert config.halo.exclusions.ticket_type_ids_search == []
        finally:
            path.unlink()

    def test_partial_exclusions(self):
        """Only some exclusion keys defined — missing ones default to []."""
        data = self._fresh_config()
        data["halo"]["exclusions"] = {"ticket_type_ids_create": [1]}
        path = write_temp_config(data)
        try:
            config = load_config(path)
            assert config.halo.exclusions.ticket_type_ids_create == [1]
            assert config.halo.exclusions.ticket_type_ids_search == []
            assert config.halo.exclusions.status_ids_search == []
        finally:
            path.unlink()
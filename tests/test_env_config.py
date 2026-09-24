import os

import pytest

from env_config import EnvConfig, config


class TestGetStr:
    def test_returns_env_value_when_set(self, monkeypatch):
        monkeypatch.setenv("SOME_STR", "hello")
        assert EnvConfig.get_str("SOME_STR") == "hello"

    def test_returns_default_when_unset(self, monkeypatch):
        monkeypatch.delenv("SOME_STR", raising=False)
        assert EnvConfig.get_str("SOME_STR", default="fallback") == "fallback"

    def test_returns_empty_string_default_when_unspecified(self, monkeypatch):
        monkeypatch.delenv("SOME_STR", raising=False)
        assert EnvConfig.get_str("SOME_STR") == ""


class TestGetInt:
    def test_returns_parsed_int_when_set(self, monkeypatch):
        monkeypatch.setenv("SOME_INT", "42")
        assert EnvConfig.get_int("SOME_INT") == 42

    def test_returns_default_when_unset(self, monkeypatch):
        monkeypatch.delenv("SOME_INT", raising=False)
        assert EnvConfig.get_int("SOME_INT", default=7) == 7

    @pytest.mark.parametrize("value", ["", "  ", "not-a-number"])
    def test_returns_default_for_invalid_values(self, monkeypatch, value):
        monkeypatch.setenv("SOME_INT", value)
        assert EnvConfig.get_int("SOME_INT", default=7) == 7


class TestGetBool:
    def test_defaults_to_false_when_unset(self, monkeypatch):
        monkeypatch.delenv("SOME_FLAG", raising=False)
        assert EnvConfig.get_bool("SOME_FLAG") is False

    def test_returns_provided_default_when_unset(self, monkeypatch):
        monkeypatch.delenv("SOME_FLAG", raising=False)
        assert EnvConfig.get_bool("SOME_FLAG", default=True) is True

    @pytest.mark.parametrize("value", ["1", "true", "True", "TRUE", "yes", "Yes"])
    def test_truthy_values_return_true(self, monkeypatch, value):
        monkeypatch.setenv("SOME_FLAG", value)
        assert EnvConfig.get_bool("SOME_FLAG") is True

    @pytest.mark.parametrize("value", ["0", "false", "False", "no", "", "  "])
    def test_falsy_values_return_false(self, monkeypatch, value):
        monkeypatch.setenv("SOME_FLAG", value)
        assert EnvConfig.get_bool("SOME_FLAG") is False


class TestKillMeNow:
    def test_reads_live_rather_than_cached(self, monkeypatch):
        monkeypatch.setenv("KILL_ME_NOW", "False")
        assert config.kill_me_now is False

        monkeypatch.setenv("KILL_ME_NOW", "True")
        assert config.kill_me_now is True


class TestSingletonSettings:
    def test_base_dir_is_a_directory(self):
        assert os.path.isdir(config.BASE_DIR)

    def test_settings_loaded_from_envtest(self):
        # conftest.py loads .envtest with override=True before this module
        # is imported, so TEST_MODE should reflect that fixture file.
        assert config.TEST_MODE is False

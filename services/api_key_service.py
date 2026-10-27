"""VWorld credentials stored only in the operating system credential vault."""

import json
from datetime import date


class ApiKeyError(Exception):
    pass


class ApiKeyService:
    SERVICE = "BioFlow.VWorld.Address"
    ACCOUNT = "api-key"

    def __init__(self, backend=None):
        self._backend = backend

    def _vault(self):
        if self._backend is not None:
            return self._backend
        try:
            import keyring
            from keyring.backends.Windows import WinVaultKeyring
            # Explicit OS backend: never silently use a plaintext plugin.
            self._backend = WinVaultKeyring()
            if self._backend.priority <= 0:
                raise RuntimeError()
            return self._backend
        except Exception:
            raise ApiKeyError("Windows 보안 저장소를 사용할 수 없습니다. keyring 설치와 접근 권한을 확인해 주세요.") from None

    def load(self):
        try:
            raw = self._vault().get_password(self.SERVICE, self.ACCOUNT)
            if raw is None:
                return "", None
            value = json.loads(raw)
            key = value["key"]
            expiry = date.fromisoformat(value["expiry"]) if value.get("expiry") else None
            if not isinstance(key, str) or not key.strip():
                raise ValueError()
            return key, expiry
        except Exception:
            raise ApiKeyError("API 키를 읽지 못했습니다. 보안 저장소를 확인하거나 키를 다시 등록해 주세요.") from None

    def save(self, key, expiry=None):
        key = key.strip()
        if not key or (expiry is not None and not isinstance(expiry, date)):
            raise ApiKeyError("API 키와 만료일 입력을 확인해 주세요.")
        try:
            self._vault().set_password(self.SERVICE, self.ACCOUNT, json.dumps({
                "key": key, "expiry": expiry.isoformat() if expiry else None,
            }))
        except Exception:
            raise ApiKeyError("API 키를 보안 저장소에 저장하지 못했습니다.") from None

    def delete(self):
        try:
            vault = self._vault()
            if vault.get_password(self.SERVICE, self.ACCOUNT) is not None:
                vault.delete_password(self.SERVICE, self.ACCOUNT)
        except Exception:
            raise ApiKeyError("API 키를 보안 저장소에서 삭제하지 못했습니다.") from None

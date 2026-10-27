"""VWorld parcel lookup. No UI dependencies or credential-bearing diagnostics."""

import json
import socket
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass

from services.coordinate_service import parse_coordinates


class GeocodingError(Exception):
    def __init__(self, code, message):
        self.code = code
        super().__init__(message)


@dataclass(frozen=True)
class VWorldAddress:
    structure: dict[str, str]
    text: str
    latitude: str
    longitude: str


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


class VWorldGeocodingService:
    ENDPOINT = "https://api.vworld.kr/req/address"

    def __init__(self, transport=None, timeout=10):
        self.transport = transport or urllib.request.build_opener(_NoRedirect()).open
        self.timeout = timeout

    def lookup(self, latitude, longitude, api_key):
        try:
            lat, lon = parse_coordinates(latitude, longitude)
        except ValueError:
            raise GeocodingError("coordinates", "유효한 위도(-90~90)와 경도(-180~180)를 입력해 주세요.") from None
        if not isinstance(api_key, str) or not api_key.strip():
            raise GeocodingError("missing_key", "VWorld API 키를 먼저 등록해 주세요.")
        query = urllib.parse.urlencode({
            "service": "address", "request": "getAddress", "version": "2.0",
            "crs": "EPSG:4326", "point": f"{lon},{lat}", "type": "PARCEL",
            "format": "json", "simple": "false", "key": api_key.strip(),
        })
        request = urllib.request.Request(f"{self.ENDPOINT}?{query}")
        try:
            with self.transport(request, timeout=self.timeout) as response:
                payload = json.loads(response.read().decode("utf-8-sig"))
        except urllib.error.HTTPError as exc:
            if exc.code in (401, 403):
                raise self._authentication_error() from None
            if exc.code == 429:
                raise GeocodingError("rate_limit", "VWorld 요청 한도에 도달했습니다. 잠시 후 다시 시도해 주세요.") from None
            raise GeocodingError("http", "VWorld 서버가 요청을 처리하지 못했습니다.") from None
        except (TimeoutError, socket.timeout):
            raise GeocodingError("timeout", "VWorld 응답 시간이 초과되었습니다.") from None
        except urllib.error.URLError as exc:
            if isinstance(exc.reason, (TimeoutError, socket.timeout)):
                raise GeocodingError("timeout", "VWorld 응답 시간이 초과되었습니다.") from None
            raise GeocodingError("network", "VWorld에 연결하지 못했습니다. 네트워크를 확인해 주세요.") from None
        except (ValueError, UnicodeError):
            raise GeocodingError("invalid_response", "VWorld JSON 응답을 읽을 수 없습니다.") from None
        except Exception:
            raise GeocodingError("network", "VWorld 요청 중 오류가 발생했습니다.") from None
        return self.parse_response(payload, lat, lon)

    @staticmethod
    def _authentication_error():
        return GeocodingError("authentication", "VWorld 인증에 실패했습니다. API 키, 허용 도메인/IP 및 유효기간을 확인해 주세요.")

    @classmethod
    def parse_response(cls, payload, latitude="", longitude=""):
        invalid = GeocodingError("invalid_response", "VWorld 주소 응답 구조가 올바르지 않습니다.")
        if not isinstance(payload, dict) or not isinstance(payload.get("response"), dict):
            raise invalid
        response = payload["response"]
        status = response.get("status")
        if status == "NOT_FOUND":
            raise GeocodingError("not_found", "해당 좌표의 지번주소가 없습니다.")
        if status == "ERROR":
            error = response.get("error")
            code = str(error.get("code", "")).upper() if isinstance(error, dict) else ""
            if code.startswith("INVALID_KEY") or code in {"EXPIRED_KEY", "UNAUTHORIZED", "PERMISSION_DENIED"}:
                raise cls._authentication_error()
            if any(token in code for token in ("LIMIT", "QUOTA", "TOO_MANY")):
                raise GeocodingError("rate_limit", "VWorld 요청 제한으로 조회하지 못했습니다. 이용 한도를 확인해 주세요.")
            raise GeocodingError("api_error", "VWorld에서 오류를 반환했습니다. 키 이용 조건과 서비스 상태를 확인해 주세요.")
        if status != "OK":
            raise invalid
        results = response.get("result")
        if not isinstance(results, list) or not results:
            raise invalid
        for result in results:
            if not isinstance(result, dict):
                raise invalid
            if result.get("type") not in (None, "parcel", "PARCEL"):
                continue
            source = result.get("structure")
            if not isinstance(source, dict):
                raise invalid
            structure = {}
            for field in ("level0", "level1", "level2", "level3", "level4L", "level4LC", "level5", "detail"):
                value = source.get(field, "")
                if value is None:
                    value = ""
                if not isinstance(value, str):
                    raise invalid
                structure[field] = value.strip()
            if not any(structure[k] for k in ("level1", "level2", "level3", "level4L", "level5", "detail")):
                raise invalid
            text = result.get("text", "")
            if not isinstance(text, str):
                raise invalid
            return VWorldAddress(structure, text, latitude, longitude)
        raise invalid

import io
import json
import unittest
import urllib.request
from unittest.mock import patch
from urllib.error import HTTPError

from course_planner.storage import (
    GraphStorage,
    SharePointStorageError,
    StorageConfigurationError,
    StorageDisabledError,
    _SafeRedirectHandler,
)


class MockResponse:
    def __init__(self, body=b""):
        self.body = body

    def read(self):
        return self.body

    def __enter__(self):
        return self

    def __exit__(self, _exc_type, _exc_value, _traceback):
        return False


def token_response(token="access-token"):
    return MockResponse(json.dumps({"access_token": token, "expires_in": 3600}).encode())


def http_error(status, headers=None, body=b""):
    return HTTPError(
        "https://not-returned-in-errors.invalid/private",
        status,
        "server detail",
        headers or {},
        io.BytesIO(body),
    )


class MockOpener:
    def open(self, request, timeout):
        return urllib.request.urlopen(request, timeout=timeout)


class GraphStorageTests(unittest.TestCase):
    def setUp(self):
        opener_patcher = patch(
            "urllib.request.build_opener", side_effect=lambda *_handlers: MockOpener()
        )
        opener_patcher.start()
        self.addCleanup(opener_patcher.stop)
        self.settings = {
            "PLANNER_SHAREPOINT_ENABLED": "true",
            "GRAPH_TENANT_ID": "tenant-id",
            "GRAPH_CLIENT_ID": "client-id",
            "GRAPH_CLIENT_SECRET": "top-secret",
            "GRAPH_DRIVE_ID": "drive-id",
            "GRAPH_INPUT_ROOT": "Inputs/Q1",
            "GRAPH_OUTPUT_ROOT": "Outputs/Q1",
        }

    def test_graph_redirect_strips_authorization_for_other_host(self):
        handler = _SafeRedirectHandler()
        request = urllib.request.Request(
            "https://graph.microsoft.com/v1.0/drives/drive-id/root:/file:/content",
            headers={"Authorization": "Bearer sensitive-token"},
        )

        redirected = handler.redirect_request(
            request,
            None,
            302,
            "Found",
            {},
            "https://download.sharepoint.example/file?sig=private",
        )

        self.assertIsNotNone(redirected)
        self.assertIsNone(redirected.get_header("Authorization"))

    def test_redirect_handler_blocks_token_redirects_and_rejects_http(self):
        request = urllib.request.Request(
            "https://login.microsoftonline.com/tenant/oauth2/v2.0/token",
            data=b"client_secret=private",
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            method="POST",
        )
        token_handler = _SafeRedirectHandler(block_redirects=True)
        self.assertIsNone(
            token_handler.redirect_request(
                request, None, 307, "Temporary Redirect", {}, "https://evil.example/token"
            )
        )

        graph_request = urllib.request.Request(
            "https://graph.microsoft.com/v1.0/file",
            headers={"Authorization": "Bearer sensitive-token"},
        )
        with self.assertRaisesRegex(SharePointStorageError, "redirect was rejected"):
            _SafeRedirectHandler().redirect_request(
                graph_request, None, 302, "Found", {}, "http://evil.example/file"
            )

    def test_token_requests_use_redirect_blocking_handler(self):
        handlers = []

        class RedirectAttemptOpener:
            def open(self, request, timeout):
                handler = handlers[-1]
                redirected = handler.redirect_request(
                    request,
                    None,
                    307,
                    "Temporary Redirect",
                    {},
                    "https://evil.example/token",
                )
                if redirected is not None:
                    raise AssertionError("token redirects must be blocked")
                raise http_error(302)

        def build_opener(*request_handlers):
            handlers.extend(request_handlers)
            return RedirectAttemptOpener()

        with patch("urllib.request.build_opener", side_effect=build_opener):
            with self.assertRaisesRegex(SharePointStorageError, "authentication failed"):
                self.make_storage().download("file.xlsx")

        self.assertEqual(len(handlers), 1)
        self.assertTrue(handlers[0].block_redirects)

    def make_storage(self, **overrides):
        settings = dict(self.settings)
        settings.update(overrides)
        return GraphStorage.from_env(settings)

    def test_disabled_by_default_does_not_make_requests(self):
        with patch("urllib.request.urlopen") as urlopen:
            storage = GraphStorage.from_env({})
            self.assertFalse(storage.enabled)
            with self.assertRaises(StorageDisabledError):
                storage.download("input.xlsx")
            with self.assertRaises(StorageDisabledError):
                storage.upload("output.xlsx", b"data")
            urlopen.assert_not_called()

    def test_enabled_configuration_requires_all_settings(self):
        for missing in self.settings:
            if missing == "PLANNER_SHAREPOINT_ENABLED":
                continue
            with self.subTest(missing=missing):
                settings = dict(self.settings)
                del settings[missing]
                with self.assertRaises(StorageConfigurationError):
                    GraphStorage.from_env(settings)

    def test_invalid_enabled_flag_is_rejected(self):
        with self.assertRaises(StorageConfigurationError):
            GraphStorage.from_env({"PLANNER_SHAREPOINT_ENABLED": "sometimes"})

    def test_download_authenticates_and_fetches_escaped_input_path(self):
        requests = []

        def fake_urlopen(request, timeout):
            requests.append(request)
            self.assertLessEqual(timeout, GraphStorage.MAX_TIMEOUT)
            if request.full_url.endswith("/token"):
                return token_response()
            return MockResponse(b"workbook")

        with patch("urllib.request.urlopen", side_effect=fake_urlopen):
            result = self.make_storage().download("Singapore/Q4 plan #1.xlsx")

        self.assertEqual(result, b"workbook")
        self.assertEqual(len(requests), 2)
        token_request, file_request = requests
        token_form = token_request.data.decode()
        self.assertIn("scope=https%3A%2F%2Fgraph.microsoft.com%2F.default", token_form)
        self.assertIn("client_credentials", token_form)
        self.assertEqual(token_request.get_method(), "POST")
        self.assertEqual(file_request.get_method(), "GET")
        self.assertEqual(file_request.get_header("Authorization"), "Bearer access-token")
        self.assertIn(
            "/Inputs/Q1/Singapore/Q4%20plan%20%231.xlsx:/content",
            file_request.full_url,
        )

    def test_upload_creates_output_hierarchy_before_put(self):
        requests = []

        def fake_urlopen(request, timeout):
            requests.append(request)
            if request.full_url.endswith("/token"):
                return token_response()
            if request.get_method() == "POST":
                name = json.loads(request.data)["name"]
                return MockResponse(
                    json.dumps({"id": f"id-{name}", "name": name, "folder": {}}).encode()
                )
            return MockResponse(b'{"id":"ignored"}')

        with patch("urllib.request.urlopen", side_effect=fake_urlopen):
            self.make_storage().upload("Singapore/result.xlsx", b"xlsx-data")

        self.assertEqual(len(requests), 5)
        create_requests = requests[1:4]
        self.assertEqual([request.get_method() for request in create_requests], ["POST"] * 3)
        first_create = create_requests[0]
        self.assertTrue(first_create.full_url.endswith("/drives/drive-id/root/children"))
        self.assertEqual(json.loads(first_create.data)["name"], "Outputs")
        self.assertEqual(json.loads(create_requests[1].data)["name"], "Q1")
        self.assertIn("/items/id-Outputs/children", create_requests[1].full_url)
        self.assertEqual(json.loads(create_requests[2].data)["name"], "Singapore")
        self.assertIn("/items/id-Q1/children", create_requests[2].full_url)

        file_request = requests[4]
        self.assertEqual(file_request.get_method(), "PUT")
        self.assertEqual(file_request.data, b"xlsx-data")
        self.assertEqual(file_request.get_header("Content-type"), "application/octet-stream")
        self.assertIn("/Outputs/Q1/Singapore/result.xlsx:/content", file_request.full_url)

    def test_ensure_folder_accepts_existing_folder_conflict(self):
        requests = []

        def fake_urlopen(request, timeout):
            requests.append(request)
            if request.full_url.endswith("/token"):
                return token_response()
            if request.get_method() == "POST":
                if json.loads(request.data)["name"] == "Q1":
                    raise http_error(409)
                return MockResponse(b'{"id":"output-id","folder":{}}')
            return MockResponse(b'{"id":"existing-q1","folder":{}}')

        with patch("urllib.request.urlopen", side_effect=fake_urlopen):
            self.make_storage().ensure_folder("")

        self.assertEqual(
            [request.get_method() for request in requests[1:]],
            ["POST", "POST", "GET"],
        )
        self.assertIn("/root:/Outputs/Q1", requests[-1].full_url)

    def test_ensure_folder_rejects_conflict_with_file(self):
        def fake_urlopen(request, timeout):
            if request.full_url.endswith("/token"):
                return token_response()
            if request.get_method() == "POST":
                if json.loads(request.data)["name"] == "Q1":
                    raise http_error(409)
                return MockResponse(b'{"id":"output-id","folder":{}}')
            return MockResponse(b'{"id":"file-id","file":{}}')

        with patch("urllib.request.urlopen", side_effect=fake_urlopen):
            with self.assertRaisesRegex(SharePointStorageError, "not a verifiable folder"):
                self.make_storage().ensure_folder("")

    def test_unsafe_paths_are_rejected_before_network(self):
        storage = self.make_storage()
        with patch("urllib.request.urlopen") as urlopen:
            for path in ("", "../secret", "a/../../secret", "/absolute", "a\\b", "a//b", "."):
                with self.subTest(path=path):
                    with self.assertRaises(SharePointStorageError):
                        storage.download(path)
            with self.assertRaises(SharePointStorageError):
                storage.ensure_folder("../secret")
            urlopen.assert_not_called()

    def test_oversized_upload_and_non_bytes_content_are_rejected(self):
        storage = self.make_storage()
        with patch("urllib.request.urlopen") as urlopen:
            with self.assertRaises(SharePointStorageError):
                storage.upload("large.bin", b"x" * (GraphStorage.MAX_UPLOAD_BYTES + 1))
            with self.assertRaises(TypeError):
                storage.upload("result.bin", bytearray(b"data"))
            urlopen.assert_not_called()

    def test_401_refreshes_token_once(self):
        calls = []

        def fake_urlopen(request, timeout):
            calls.append(request)
            if request.full_url.endswith("/token"):
                token_number = sum(item.full_url.endswith("/token") for item in calls)
                return token_response(f"token-{token_number}")
            if request.get_header("Authorization") == "Bearer token-1":
                raise http_error(401, body=b"do not expose this")
            return MockResponse(b"file")

        with patch("urllib.request.urlopen", side_effect=fake_urlopen):
            result = self.make_storage().download("item.xlsx")

        self.assertEqual(result, b"file")
        self.assertEqual(sum(call.full_url.endswith("/token") for call in calls), 2)
        graph_requests = [call for call in calls if not call.full_url.endswith("/token")]
        self.assertEqual(
            [call.get_header("Authorization") for call in graph_requests],
            ["Bearer token-1", "Bearer token-2"],
        )

    def test_429_retries_with_bounded_retry_after(self):
        calls = []

        def fake_urlopen(request, timeout):
            calls.append(request)
            if request.full_url.endswith("/token"):
                return token_response()
            if len([call for call in calls if not call.full_url.endswith("/token")]) == 1:
                raise http_error(429, {"Retry-After": "100"})
            return MockResponse(b"file")

        with patch("urllib.request.urlopen", side_effect=fake_urlopen) as urlopen:
            with patch("time.sleep") as sleep:
                self.assertEqual(self.make_storage().download("item.xlsx"), b"file")

        sleep.assert_called_once_with(GraphStorage.MAX_RETRY_DELAY)
        self.assertEqual(urlopen.call_count, 3)

    def test_retries_are_bounded_and_errors_are_sanitized(self):
        def fake_urlopen(request, timeout):
            if request.full_url.endswith("/token"):
                return token_response()
            raise http_error(
                503,
                {"Retry-After": "0"},
                b"top-secret https://private.invalid/token",
            )

        with patch("urllib.request.urlopen", side_effect=fake_urlopen) as urlopen:
            with patch("time.sleep"):
                with self.assertRaises(SharePointStorageError) as raised:
                    self.make_storage().download("item.xlsx")

        self.assertEqual(urlopen.call_count, 2 + GraphStorage.DEFAULT_RETRIES)
        message = str(raised.exception)
        self.assertIn("503", message)
        self.assertNotIn("top-secret", message)
        self.assertNotIn("private.invalid", message)

    def test_authentication_errors_do_not_expose_secret_or_url(self):
        with patch(
            "urllib.request.urlopen",
            side_effect=http_error(
                400,
                body=b"top-secret https://private.invalid/oauth2/token",
            ),
        ):
            with self.assertRaises(SharePointStorageError) as raised:
                self.make_storage().download("item.xlsx")

        message = str(raised.exception)
        self.assertNotIn("top-secret", message)
        self.assertNotIn("private.invalid", message)
        self.assertIn("400", message)


if __name__ == "__main__":
    unittest.main()

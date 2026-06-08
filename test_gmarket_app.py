import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import gmarket
import gmarket_app


class GmarketAppTests(unittest.TestCase):
    def test_validate_run_time_accepts_hh_mm(self):
        self.assertEqual(gmarket_app.validate_run_time("09:05"), "09:05")
        self.assertEqual(gmarket_app.validate_run_time("23:59"), "23:59")

    def test_validate_run_time_rejects_bad_values(self):
        for value in ["9:05", "24:00", "12:60", "abc", ""]:
            with self.subTest(value=value):
                with self.assertRaises(ValueError):
                    gmarket_app.validate_run_time(value)

    def test_save_and_load_config_roundtrip(self):
        with tempfile.TemporaryDirectory() as tmp:
            config_path = Path(tmp) / "config.json"
            config = {
                "workbook_path": r"C:\data\빅스_판매량_20260604.xlsx",
                "run_time": "14:00",
            }

            gmarket_app.save_config(config, config_path)

            self.assertEqual(gmarket_app.load_config(config_path), config)
            self.assertEqual(
                json.loads(config_path.read_text(encoding="utf-8")),
                config,
            )

    def test_build_schtasks_args_registers_daily_run_command(self):
        args = gmarket_app.build_schtasks_args(
            run_command=[r"C:\Program Files\GmarketCrawler\GmarketCrawler.exe", "--run"],
            run_time="14:30",
        )

        self.assertEqual(args[:3], ["schtasks", "/Create", "/TN"])
        self.assertIn(gmarket_app.TASK_NAME, args)
        self.assertIn("/SC", args)
        self.assertIn("DAILY", args)
        self.assertIn("/ST", args)
        self.assertIn("14:30", args)
        self.assertIn("/F", args)
        self.assertIn('"C:\\Program Files\\GmarketCrawler\\GmarketCrawler.exe" --run', args)

    def test_source_run_command_uses_python_script(self):
        with patch.object(sys, "frozen", False, create=True):
            command = gmarket_app.get_run_command()

        self.assertEqual(Path(command[0]), Path(sys.executable))
        self.assertEqual(Path(command[1]).name, "gmarket_app.py")
        self.assertEqual(command[2], "--run")

    def test_scheduled_run_logs_missing_workbook_failure(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            config_path = tmp_path / "config.json"
            log_path = tmp_path / "run.log"
            gmarket_app.save_config(
                {
                    "workbook_path": str(tmp_path / "missing.xlsx"),
                    "run_time": "14:00",
                },
                config_path,
            )

            with patch.object(gmarket_app, "get_log_path", return_value=log_path):
                with self.assertRaises(ValueError):
                    gmarket_app.run_crawler_from_config(config_path)

            log_text = log_path.read_text(encoding="utf-8")
            self.assertIn("실행 실패", log_text)
            self.assertIn("엑셀 파일을 찾을 수 없습니다", log_text)

    def test_gmarket_load_products_accepts_workbook_path(self):
        self.assertTrue(callable(gmarket.load_products_from_workbook))
        self.assertIn("workbook_path", gmarket.load_products_from_workbook.__code__.co_varnames)

    def test_parse_chrome_major_version(self):
        self.assertEqual(gmarket.parse_chrome_major_version("Google Chrome 146.0.7390.122"), 146)
        self.assertEqual(gmarket.parse_chrome_major_version("147.0.1"), 147)
        self.assertIsNone(gmarket.parse_chrome_major_version("not chrome"))

    def test_get_chrome_installation_uses_file_version_when_version_output_missing(self):
        with tempfile.TemporaryDirectory() as tmp:
            chrome_path = Path(tmp) / "chrome.exe"
            chrome_path.write_text("", encoding="utf-8")

            with patch.object(
                gmarket,
                "get_chrome_executable_candidates",
                return_value=[chrome_path],
            ):
                with patch.object(
                    gmarket.subprocess,
                    "run",
                    return_value=Mock(stdout="기존 브라우저 세션에서 여는 중입니다.", stderr=""),
                ):
                    with patch.object(
                        gmarket,
                        "get_windows_file_major_version",
                        return_value=148,
                    ):
                        self.assertEqual(
                            gmarket.get_chrome_installation(),
                            (148, chrome_path),
                        )

    def test_create_chrome_driver_uses_automatic_version_first(self):
        driver = object()

        with patch.object(gmarket, "get_chrome_installation", return_value=None):
            with patch.object(gmarket.uc, "Chrome", return_value=driver) as chrome:
                self.assertIs(gmarket.create_chrome_driver(Mock()), driver)

        self.assertNotIn("version_main", chrome.call_args.kwargs)

    def test_create_chrome_driver_uses_detected_chrome_first(self):
        driver = object()
        chrome_path = Path(r"C:\Chrome148\chrome.exe")

        with patch.object(
            gmarket,
            "get_chrome_installation",
            return_value=(148, chrome_path),
        ):
            with patch.object(gmarket.uc, "Chrome", return_value=driver) as chrome:
                self.assertIs(gmarket.create_chrome_driver(Mock()), driver)

        self.assertEqual(chrome.call_count, 1)
        self.assertEqual(chrome.call_args.kwargs["version_main"], 148)
        self.assertEqual(
            chrome.call_args.kwargs["browser_executable_path"],
            str(chrome_path),
        )

    def test_create_chrome_driver_raises_original_error_when_chrome_not_detected(self):
        with patch.object(gmarket, "get_chrome_installation", return_value=None):
            with patch.object(
                gmarket.uc,
                "Chrome",
                side_effect=RuntimeError("driver mismatch"),
            ):
                with self.assertRaisesRegex(RuntimeError, "driver mismatch"):
                    gmarket.create_chrome_driver(Mock())

    def test_create_chrome_driver_does_not_reuse_options_after_auto_failure(self):
        driver = object()
        chrome_path = Path(r"C:\Chrome148\chrome.exe")

        with patch.object(
            gmarket,
            "get_chrome_installation",
            return_value=(148, chrome_path),
        ):
            with patch.object(gmarket.uc, "Chrome", return_value=driver) as chrome:
                self.assertIs(gmarket.create_chrome_driver(Mock()), driver)

        self.assertEqual(chrome.call_count, 1)

    def test_create_chrome_driver_uses_detected_major_version(self):
        driver = object()

        with patch.object(
            gmarket.uc,
            "Chrome",
            return_value=driver,
        ) as chrome:
            with patch.object(
                gmarket,
                "get_chrome_installation",
                return_value=(146, Path(r"C:\Chrome146\chrome.exe")),
            ):
                self.assertIs(gmarket.create_chrome_driver(Mock()), driver)

        self.assertEqual(chrome.call_args.kwargs["version_main"], 146)

    def test_create_chrome_driver_uses_detected_chrome_binary(self):
        driver = object()

        with tempfile.TemporaryDirectory() as tmp:
            chrome_path = Path(tmp) / "chrome.exe"
            chrome_path.write_text("", encoding="utf-8")

            with patch.object(
                gmarket.uc,
                "Chrome",
                return_value=driver,
            ) as chrome:
                with patch.object(
                    gmarket,
                    "get_chrome_executable_candidates",
                    return_value=[chrome_path],
                ):
                    with patch.object(
                        gmarket.subprocess,
                        "run",
                        return_value=Mock(stdout="Google Chrome 148.0.7778.181", stderr=""),
                    ):
                        self.assertIs(gmarket.create_chrome_driver(Mock()), driver)

        self.assertEqual(chrome.call_args.kwargs["version_main"], 148)
        self.assertEqual(
            chrome.call_args.kwargs["browser_executable_path"],
            str(chrome_path),
        )

    def test_configure_ssl_certificates_uses_truststore(self):
        truststore = Mock()

        with patch.object(gmarket.importlib, "import_module", return_value=truststore):
            result = gmarket.configure_ssl_certificates()

        self.assertEqual(result, "truststore")
        truststore.inject_into_ssl.assert_called_once_with()

    def test_configure_ssl_certificates_falls_back_to_certifi(self):
        certifi = Mock()
        certifi.where.return_value = r"C:\certs\cacert.pem"

        def import_module(name):
            if name == "truststore":
                raise ImportError("missing truststore")
            if name == "certifi":
                return certifi
            raise AssertionError(name)

        with patch.object(gmarket.importlib, "import_module", side_effect=import_module):
            with patch.dict(gmarket.os.environ, {}, clear=True):
                result = gmarket.configure_ssl_certificates()

                self.assertEqual(result, "certifi")
                self.assertEqual(
                    gmarket.os.environ["SSL_CERT_FILE"],
                    r"C:\certs\cacert.pem",
                )

    def test_create_chrome_driver_configures_ssl_before_uc_chrome(self):
        driver = object()

        with patch.object(gmarket, "get_chrome_installation", return_value=None):
            with patch.object(
                gmarket,
                "configure_ssl_certificates",
                return_value="truststore",
            ) as configure_ssl:
                with patch.object(gmarket.uc, "Chrome", return_value=driver):
                    self.assertIs(gmarket.create_chrome_driver(Mock()), driver)

        configure_ssl.assert_called_once_with()

if __name__ == "__main__":
    unittest.main()

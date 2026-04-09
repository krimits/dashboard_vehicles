import importlib.util
import json
import os
import tempfile
import threading
import unittest
import urllib.error
import urllib.request
from pathlib import Path
from urllib.parse import quote


WORKSPACE = Path(__file__).parent
MODULE_PATH = WORKSPACE / "Qwen_python_20260317_kqvga2wu9.py"
EXCEL_PATH = WORKSPACE / "ΔΕΛΤΙΟ ΕΝΗΜΕΡΩΣΗΣ #001 2026-03-17 ΗΜΕΡΗΣΙΑ ΚΑΤΑΣΤΑΣΗ ΟΧΗΜΑΤΩΝ.xls"


def load_module():
    spec = importlib.util.spec_from_file_location("fleet_dashboard_app", MODULE_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


class FindLatestExcelFileTests(unittest.TestCase):
    def test_ignores_lock_files_and_picks_newest_matching_excel(self):
        module = load_module()

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            older = temp_path / "ΔΕΛΤΙΟ ΕΝΗΜΕΡΩΣΗΣ #001 2026-03-16 ΗΜΕΡΗΣΙΑ ΚΑΤΑΣΤΑΣΗ ΟΧΗΜΑΤΩΝ.xls"
            newer = temp_path / "ΔΕΛΤΙΟ ΕΝΗΜΕΡΩΣΗΣ #001 2026-03-17 ΗΜΕΡΗΣΙΑ ΚΑΤΑΣΤΑΣΗ ΟΧΗΜΑΤΩΝ.xls"
            lock_file = temp_path / "~$ΔΕΛΤΙΟ ΕΝΗΜΕΡΩΣΗΣ #001 2026-03-18 ΗΜΕΡΗΣΙΑ ΚΑΤΑΣΤΑΣΗ ΟΧΗΜΑΤΩΝ.xlsx"

            older.write_text("old", encoding="utf-8")
            newer.write_text("new", encoding="utf-8")
            lock_file.write_text("lock", encoding="utf-8")

            older.touch()
            newer.touch()
            lock_file.touch()

            result = module.find_latest_excel_file(temp_path)

            self.assertEqual(result, newer)


class CollectionDailyAvailabilityParserTests(unittest.TestCase):
    def test_parses_rows_after_header_until_total(self):
        module = load_module()
        # Παλιά διάταξη: κεφαλίδα μετά τη ζώνη 5–26, δύο στήλες (fallback blob).
        prefix = [[None, None] for _ in range(28)]
        rows = prefix + [
            [
                "ΟΧΗΜΑΤΑ ΑΠΟΚΟΜΙΔΗΣ (ΗΜΕΡΗΣΙΑ ΔΙΑΘΕΣΙΜΟΤΗΤΑ)",
                None,
            ],
            ["ΜΙΚΡΑ 4Τ (ΜΥΛΟΙ) (IVECO)", 5],
            ["ΝΕΑ ΜΙΚΡΑ 2Τ (ΠΡΕΣΣΑΚΙΑ)", 3],
            ["ΣΥΝΟΛΟ (TOTAL)", 65],
        ]
        result = module.parse_collection_daily_availability(rows)
        self.assertEqual(
            result,
            [
                {"name": "ΜΙΚΡΑ 4Τ (ΜΥΛΟΙ) (IVECO)", "count": 5},
                {"name": "ΝΕΑ ΜΙΚΡΑ 2Τ (ΠΡΕΣΣΑΚΙΑ)", "count": 3},
            ],
        )

    def test_parses_table_at_column_l_same_row_band_as_main_22(self):
        """Τυπικό δελτίο: L17:M28 — κεφαλίδα στη στήλη L ενώ οι γραμμές 5–26 είναι το κύριο summary."""
        module = load_module()
        wide = 13
        rows = [[None] * wide for _ in range(16)]
        rows.append(
            [None] * 11
            + [
                "ΟΧΗΜΑΤΑ ΑΠΟΚΟΜΙΔΗΣ (ΗΜΕΡΗΣΙΑ ΔΙΑΘΕΣΙΜΟΤΗΤΑ)",
                None,
            ]
        )
        rows.append([None] * 11 + ["ΜΙΚΡΑ 4Τ (ΜΥΛΟΙ) (IVECO)", 5])
        rows.append([None] * 11 + ["ΝΕΑ ΜΙΚΡΑ 2Τ (ΠΡΕΣΣΑΚΙΑ)", 3])
        rows.append([None] * 11 + ["ΣΥΝΟΛΟ (TOTAL)", 65])
        result = module.parse_collection_daily_availability(rows)
        self.assertEqual(
            result,
            [
                {"name": "ΜΙΚΡΑ 4Τ (ΜΥΛΟΙ) (IVECO)", "count": 5},
                {"name": "ΝΕΑ ΜΙΚΡΑ 2Τ (ΠΡΕΣΣΑΚΙΑ)", "count": 3},
            ],
        )

    def test_returns_empty_when_header_missing(self):
        module = load_module()
        rows = [["Άλλο κείμενο", 1], ["Κατηγορία", 2]]
        self.assertEqual(module.parse_collection_daily_availability(rows), [])

    def test_ignores_orphan_header_without_data_rows(self):
        module = load_module()
        rows = [
            [
                "ΟΧΗΜΑΤΑ ΑΠΟΚΟΜΙΔΗΣ (ΗΜΕΡΗΣΙΑ ΔΙΑΘΕΣΙΜΟΤΗΤΑ)",
                None,
            ],
            ["ΨΕΥΔΟ-ΔΕΔΟΜΕΝΑ", 99],
        ]
        self.assertEqual(module.parse_collection_daily_availability(rows), [])

    def test_uses_header_column_band_not_left_summary_columns(self):
        """Στις ίδιες γραμμές το αριστερό block είναι το summary 22· ο πίνακας αποκομιδής είναι δεξιά."""
        module = load_module()
        wide = 14
        rows = [[None] * wide for _ in range(16)]
        rows.append(
            [None] * 11
            + [
                "ΟΧΗΜΑΤΑ ΑΠΟΚΟΜΙΔΗΣ (ΗΜΕΡΗΣΙΑ ΔΙΑΘΕΣΙΜΟΤΗΤΑ)",
                None,
            ]
        )
        rows.append(
            [None, "ΛΕΩΦΟΡΕΙΑ ΣΧΟΛΙΚΑ", None, None, 12, None, None, None, None, None, None, "ΜΙΚΡΑ 4Τ (ΜΥΛΟΙ) (IVECO)", 6]
        )
        rows.append(
            [None, "ΑΛΛΗ ΚΑΤΗΓΟΡΙΑ", None, None, 5, None, None, None, None, None, None, "ΝΕΑ ΜΙΚΡΑ 2Τ (ΠΡΕΣΣΑΚΙΑ)", 5]
        )
        rows.append([None] * 11 + ["ΣΥΝΟΛΟ (TOTAL)", 66])
        result = module.parse_collection_daily_availability(rows)
        self.assertEqual(
            result,
            [
                {"name": "ΜΙΚΡΑ 4Τ (ΜΥΛΟΙ) (IVECO)", "count": 6},
                {"name": "ΝΕΑ ΜΙΚΡΑ 2Τ (ΠΡΕΣΣΑΚΙΑ)", "count": 5},
            ],
        )


class SummaryKpiTripletTests(unittest.TestCase):
    def test_infers_total_when_summary_cell_empty_but_ins_brk_present(self):
        module = load_module()
        rows = [[None] * 14 for _ in range(20)]
        rows[12] = [None] * 13 + [59]
        rows[13] = [None] * 13 + [40]
        rows[14] = [None] * 13 + [0]
        cats = [{"total": 1, "in_service": 1, "broken": 0} for _ in range(22)]
        ins, brk, tot = module._parse_summary_kpi_triplet(rows, 12, 13, 14, 13, cats)
        self.assertEqual((ins, brk, tot), (59, 40, 99))


class WorkbookParsingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.module = load_module()

    def test_extracts_total_and_real_summary_from_workbook(self):
        payload = self.module.load_dashboard_payload(EXCEL_PATH)

        self.assertEqual(payload["summary"]["all_vehicles"]["total"], 232)
        self.assertEqual(payload["summary"]["all_vehicles"]["in_service"], 95)
        self.assertEqual(payload["summary"]["all_vehicles"]["broken"], 137)
        self.assertEqual(payload["summary"]["real_fleet"]["total"], 164)
        self.assertEqual(payload["summary"]["real_fleet"]["in_service"], 95)
        self.assertEqual(payload["summary"]["real_fleet"]["broken"], 69)
        self.assertIn("collection_daily_availability", payload["summary"])
        self.assertIsInstance(payload["summary"]["collection_daily_availability"], list)

    def test_builds_category_drilldown_with_vehicle_rows(self):
        payload = self.module.load_dashboard_payload(EXCEL_PATH)

        category = next(
            item for item in payload["categories"]
            if item["name"] == "ΠΡΕΣΕΣ 5Τ ΝΕΕΣ"
        )

        self.assertEqual(category["summary"]["total"], 11)
        self.assertEqual(category["summary"]["in_service"], 4)
        self.assertEqual(category["summary"]["broken"], 7)
        self.assertGreaterEqual(len(category["vehicles"]), 11)
        self.assertTrue(
            any(vehicle["vehicle_id"] == "825" for vehicle in category["vehicles"])
        )

    def test_marks_explicitly_or_inferred_excluded_vehicles(self):
        payload = self.module.load_dashboard_payload(EXCEL_PATH)

        category = next(
            item for item in payload["categories"]
            if item["name"] == "ΠΑΛΙΑ ΜΙΚΡΑ 2Τ (ΜΥΛΟΙ)"
        )

        excluded = [vehicle for vehicle in category["vehicles"] if vehicle["excluded"]]
        self.assertEqual(len(excluded), 7)
        self.assertTrue(all(vehicle["exclusion_reason"] for vehicle in excluded))


class DashboardHtmlTests(unittest.TestCase):
    def test_dashboard_html_uses_live_api_and_tree_regions(self):
        module = load_module()

        html = module.build_dashboard_html()

        self.assertIn("/api/fleet-data", html)
        self.assertIn("availabilityTree", html)
        self.assertIn("detailsTree", html)
        self.assertIn("setInterval(loadDashboardData", html)

    def test_dashboard_html_contains_comparative_tab(self):
        module = load_module()

        html = module.build_dashboard_html()

        self.assertIn('id="comparative"', html)
        self.assertIn("Συγκριτικά Στατιστικά στοιχεία", html)
        self.assertIn("refreshLiveComparativeUI", html)
        self.assertIn("live_fleet_comparative_baseline", html)

    def test_dashboard_html_contains_overview_bar_chart_containers(self):
        module = load_module()

        html = module.build_dashboard_html()

        self.assertIn('id="overviewRealFleetBarChart"', html)
        self.assertIn('id="overviewCollectionBarChart"', html)
        self.assertIn("overviewRealFleetBarChart", html)
        self.assertIn("collection_daily_availability", html)

    def test_dashboard_html_and_messages_do_not_contain_garbled_greek(self):
        module = load_module()

        html = module.build_dashboard_html()
        not_found = module.DashboardRequestHandler.send_error.__defaults__

        self.assertNotIn("?�", html)
        self.assertNotIn("\ufffd", html)
        self.assertNotIn("?�", module.clean_text(module.EXCEL_PREFIX) or "")

    def test_dashboard_html_contains_admin_entry_point(self):
        module = load_module()

        html = module.build_dashboard_html()

        self.assertIn("/admin", html)


class UploadedWorkbookTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.module = load_module()

    def test_uploaded_workbook_becomes_active_source(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            base_dir = Path(temp_dir)
            storage = self.module.AppStorage(base_dir)

            uploaded_path = storage.activate_uploaded_workbook(
                filename=EXCEL_PATH.name,
                file_bytes=EXCEL_PATH.read_bytes(),
            )

            self.assertTrue(uploaded_path.exists())
            self.assertEqual(storage.resolve_active_workbook(), uploaded_path)

    def test_upload_rejects_invalid_extension(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            storage = self.module.AppStorage(Path(temp_dir))

            with self.assertRaises(self.module.DashboardDataError):
                storage.activate_uploaded_workbook(
                    filename="bad.txt",
                    file_bytes=b"not-an-excel-file",
                )

    def test_uploaded_workbook_is_preferred_over_directory_scan(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            base_dir = Path(temp_dir)
            storage = self.module.AppStorage(base_dir)

            fallback = base_dir / EXCEL_PATH.name
            fallback.write_bytes(EXCEL_PATH.read_bytes())

            uploaded_path = storage.activate_uploaded_workbook(
                filename=EXCEL_PATH.name,
                file_bytes=EXCEL_PATH.read_bytes(),
            )

            resolved = self.module.resolve_workbook_path(base_dir, storage)
            self.assertEqual(resolved, uploaded_path)


class AdminSecurityTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.module = load_module()

    def test_admin_secret_check_works(self):
        self.assertTrue(self.module.is_authorized_admin("secret", "secret"))
        self.assertFalse(self.module.is_authorized_admin("secret", "wrong"))
        self.assertFalse(self.module.is_authorized_admin("", "wrong"))


class OnlineWorkflowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.module = load_module()

    def test_public_dashboard_and_admin_upload_workflow(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            base_dir = Path(temp_dir)
            server = self.module.build_server(
                host="127.0.0.1",
                port=0,
                base_directory=base_dir,
                storage_dir=base_dir / "state",
                admin_secret="secret",
            )
            thread = threading.Thread(target=server.serve_forever, daemon=True)
            thread.start()
            base_url = f"http://127.0.0.1:{server.server_address[1]}"

            try:
                with urllib.request.urlopen(f"{base_url}/") as response:
                    html = response.read().decode("utf-8")
                self.assertIn("/api/fleet-data", html)
                self.assertIn("/admin", html)

                unauthorized_request = urllib.request.Request(
                    f"{base_url}/api/admin/upload",
                    data=EXCEL_PATH.read_bytes(),
                    method="POST",
                    headers={
                        "Content-Type": "application/octet-stream",
                        "X-Upload-Filename": quote(EXCEL_PATH.name),
                    },
                )
                with self.assertRaises(urllib.error.HTTPError) as unauthorized_error:
                    urllib.request.urlopen(unauthorized_request)

                self.assertEqual(unauthorized_error.exception.code, 401)
                unauthorized_payload = json.loads(unauthorized_error.exception.read().decode("utf-8"))
                self.assertIn("Μη εξουσιοδοτημένο", unauthorized_payload["error"]["message"])

                upload_request = urllib.request.Request(
                    f"{base_url}/api/admin/upload",
                    data=EXCEL_PATH.read_bytes(),
                    method="POST",
                    headers={
                        "Content-Type": "application/octet-stream",
                        "X-Upload-Filename": quote(EXCEL_PATH.name),
                        "X-Admin-Secret": "secret",
                    },
                )
                with urllib.request.urlopen(upload_request) as response:
                    self.assertEqual(response.status, 201)
                    upload_payload = json.loads(response.read().decode("utf-8"))

                active_file = upload_payload["upload"]["active_file"]
                self.assertTrue(active_file.endswith(".xls"))

                with urllib.request.urlopen(f"{base_url}/api/fleet-data") as response:
                    payload = json.loads(response.read().decode("utf-8"))

                self.assertEqual(payload["summary"]["real_fleet"]["total"], 164)
                self.assertEqual(payload["meta"]["source_file"], active_file)
                self.assertIn(":", payload["meta"]["version"])
            finally:
                server.shutdown()
                server.server_close()
                thread.join(timeout=2)


if __name__ == "__main__":
    unittest.main()

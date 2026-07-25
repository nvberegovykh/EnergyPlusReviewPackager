import importlib.util
import os
import tempfile
import unittest

from bs4 import BeautifulSoup


MODULE_PATH = os.path.join(os.path.dirname(__file__), "EnergyPlusReviewPackager.py")
SPEC = importlib.util.spec_from_file_location("report_packager", MODULE_PATH)
PACKAGER = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(PACKAGER)


class ScheduleProfileTests(unittest.TestCase):
    def test_two_html_summary_samples_are_not_a_daily_curve(self):
        html = """
        <html><body><b>Schedules-SetPoints</b><table>
        <tr><th>First Object Used</th><th>11am [C]</th><th>11pm [C]</th></tr>
        <tr><td>Thermostat 1</td><td>21</td><td>13</td></tr>
        </table></body></html>
        """
        profiles = PACKAGER.extract_schedule_profiles(
            BeautifulSoup(html, "html.parser")
        )
        self.assertEqual([], profiles)

    def test_complete_24_value_html_curve_is_accepted(self):
        headers = "".join(f"<th>{hour}</th>" for hour in range(1, 25))
        values = "".join(f"<td>{hour}</td>" for hour in range(1, 25))
        html = (
            "<html><body><b>Schedules-SetPoints</b><table>"
            f"<tr><th>Schedule</th>{headers}</tr>"
            f"<tr><td>Daily</td>{values}</tr>"
            "</table></body></html>"
        )
        profiles = PACKAGER.extract_schedule_profiles(
            BeautifulSoup(html, "html.parser")
        )
        self.assertEqual(1, len(profiles))
        self.assertEqual(24, len(profiles[0]["vals"]))

    def test_in_idf_is_auto_discovered_beside_report(self):
        with tempfile.TemporaryDirectory() as folder:
            report = os.path.join(folder, "eplustbl.html")
            model = os.path.join(folder, "in.idf")
            with open(report, "w", encoding="utf-8") as handle:
                handle.write("<html></html>")
            with open(model, "w", encoding="utf-8") as handle:
                handle.write("Version, 24.1;")
            self.assertEqual(model, PACKAGER.discover_schedule_model(report))

    def test_current_energyplus_run_has_real_daily_profiles(self):
        root = os.path.dirname(os.path.dirname(__file__))
        report = os.path.join(root, "current_run", "eplustbl.htm")
        model = os.path.join(root, "current_run", "in.idf")
        self.assertEqual(model, PACKAGER.discover_schedule_model(report))
        soup = PACKAGER.load_soup(report)
        used = PACKAGER.extract_used_schedule_names(soup)
        refs = PACKAGER.extract_used_schedule_refs_from_model(model)
        profiles = PACKAGER.extract_model_schedule_profiles(
            model,
            used_schedule_names=set(used) | set(refs.get("names", set())),
            used_schedule_handles=refs.get("handles", set()),
            max_profiles=120,
        )
        self.assertGreater(len(profiles), 0)
        self.assertTrue(all(len(profile["vals"]) == 24 for profile in profiles))
        names = {profile["name"] for profile in profiles}
        self.assertIn("Interior_Lighting_Schedule_0.0000", names)
        self.assertIn("Activity_Always_120W", names)
        self.assertIn("Heating Setpoint", names)
        self.assertIn("Cooling Setpoint", names)
        self.assertFalse(any(name.startswith("Schedule Day ") for name in names))
        heating = next(profile for profile in profiles if profile["name"] == "Heating Setpoint")
        self.assertEqual("F", heating["unit"])
        self.assertAlmostEqual(55.0811, heating["vals"][0], places=3)
        self.assertAlmostEqual(69.9980, heating["vals"][6], places=3)


if __name__ == "__main__":
    unittest.main()

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import _apply_single_supplemental_legend


def test_single_legend_image_fills_schedule_image_without_legend():
    sheets = [
        {
            "filename": "田村さんスケジュール.jpg",
            "legend": [],
            "employees": [{"name": "田村", "shifts": [{"date": "2026-05-01", "code": "2"}]}],
            "off_markers": [""],
        },
        {
            "filename": "田村シフト.jpeg",
            "legend": [
                {
                    "code": "2",
                    "label": "2番勤務",
                    "start_time": "08:30",
                    "end_time": "17:00",
                    "is_off": False,
                },
            ],
            "employees": [],
            "off_markers": ["夏"],
        },
    ]

    merged = _apply_single_supplemental_legend(sheets)

    assert len(merged) == 1
    assert merged[0]["filename"] == "田村さんスケジュール.jpg"
    assert merged[0]["legend"][0]["code"] == "2"
    assert merged[0]["off_markers"] == ["", "夏"]

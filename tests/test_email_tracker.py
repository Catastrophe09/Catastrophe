from src.email_tracker import build_task_from_email


def test_build_task_from_email_maps_fields() -> None:
    message = {
        "subject": "Quarterly report",
        "from": {"emailAddress": {"address": "boss@example.com"}},
        "receivedDateTime": "2026-03-01T10:15:00Z",
        "webLink": "https://outlook.office.com/mail/id/test",
    }

    task = build_task_from_email(message)

    assert task["title"] == "Quarterly report"
    assert "boss@example.com" in task["body"]["content"]
    assert "2026-03-01T10:15:00Z" in task["body"]["content"]
    assert "https://outlook.office.com/mail/id/test" in task["body"]["content"]


def test_build_task_from_email_handles_missing_fields() -> None:
    task = build_task_from_email({})

    assert task["title"] == "(No subject)"
    assert "unknown" in task["body"]["content"]

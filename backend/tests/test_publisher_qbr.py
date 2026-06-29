import unittest

from app.services.publisher_qbr import build_publisher_qbr_record


class PublisherQbrRecordTests(unittest.TestCase):
    def test_builds_record_from_extension_payload(self):
        record = build_publisher_qbr_record(
            {
                "type": "PUBLISHER_QBR_REQUEST",
                "sourceID": "primary-site",
                "fromDate": "20260101",
                "toDate": "20260331",
                "currencyCode": "gbp",
                "td_tokens": {"impersonate_access_token": "secret-token"},
                "comparisonPublishers": [
                    {"sourceID": "site-1"},
                    {"siteID": "site-2"},
                    {"sourceId": "site-3"},
                    {"primarySourceId": "site-4"},
                    {"sourceID": "site-5"},
                ],
            }
        )

        self.assertIsNotNone(record)
        assert record is not None
        self.assertEqual(record["primary_publisher_site_id"], "primary-site")
        self.assertEqual(record["comparison_publisher_site_ids"], ["site-1", "site-2", "site-3", "site-4"])
        self.assertEqual(record["from_date"], "2026-01-01")
        self.assertEqual(record["to_date"], "2026-03-31")
        self.assertEqual(record["currency_code"], "GBP")
        self.assertEqual(record["request_payload"]["td_tokens"], "[redacted]")

    def test_ignores_non_publisher_payload_without_site_id(self):
        self.assertIsNone(
            build_publisher_qbr_record(
                {
                    "type": "QBR_REQUEST",
                    "programId": "123",
                    "fromDate": "20260101",
                    "toDate": "20260331",
                    "currencyCode": "EUR",
                }
            )
        )


if __name__ == "__main__":
    unittest.main()

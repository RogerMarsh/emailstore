# test_emailcollector.py
# Copyright 2014 Roger Marsh
# Licence: See LICENCE (BSD licence)

"""emailcollector tests."""

import unittest
import os

from .. import emailcollector


class EmailCollector(unittest.TestCase):
    def setUp(self):
        pass

    def tearDown(self):
        pass

    def test__assumptions(self):
        msg = "Failure of this test invalidates all other tests"
        self.assertEqual(emailcollector._OPERA_EMAIL_CLIENT, "opera", msg=msg)

    def test___init__01(self):
        ec = emailcollector.EmailCollector("d")
        self.assertEqual(ec.directory, "d")
        self.assertEqual(ec.configuration, None)
        self.assertEqual(ec.dryrun, True)
        self.assertEqual(ec.parent, None)
        self.assertEqual(ec.criteria, None)
        self.assertEqual(ec.email_client, None)
        self.assertEqual(len(ec.__dict__), 6)

    def test_parse_01(self):
        ec = emailcollector.EmailCollector(
            "d",
            configuration="\n".join(
                (
                    "collected collected",
                    "mailboxstyle mbox",
                    "mboxmailstore ~/test.mbs",
                    "earliestfromdate 2014-02-03",
                    "mostrecentfromdate 2014-06-20",
                    "account r.m@rmswch.plus.com",
                    "account roger.marsh@btinternet.com",
                    "emailsfrom mcclarkeeastl@yahoo.co.uk",
                )
            ),
        )
        self.assertEqual(ec.parse(), True)


class EmailCollector_select(unittest.TestCase):
    def setUp(self):
        self.opd = os.path.join("~", "testoperaselect")
        try:
            for f in os.listdir(os.path.expanduser(self.opd)):
                try:
                    os.remove(os.path.expanduser(os.path.join(self.opd, f)))
                except OSError:
                    pass
            try:
                os.rmdir(os.path.expanduser(self.opd))
            except OSError:
                pass
        except OSError:
            pass

    def tearDown(self):
        pass

    def test__select_emails_01(self):
        outputdirectory = " ".join(
            (
                "outputdirectory",
                os.path.expanduser(self.opd),
            )
        )
        ec = emailcollector.EmailCollector(
            "d",
            configuration="\n".join(
                (
                    "collected collected",
                    "mailboxstyle mbox",
                    "mboxmailstore ~/test.mbs",
                    "earliestfromdate 2014-02-03",
                    "mostrecentfromdate 2014-06-20",
                    "account r.m@rmswch.plus.com",
                    "account roger.marsh@btinternet.com",
                    "emailsfrom mcclarkeeastl@yahoo.co.uk",
                )
            ),
        )
        self.assertEqual(ec.parse(), True)
        self.assertEqual(ec.email_client, None)
        self.assertEqual(ec.selected_emails, ec.selected_emails)
        self.assertEqual(ec.selected_emails_text, ec.selected_emails_text)


def suite_ec():
    return unittest.TestLoader().loadTestsFromTestCase(EmailCollector)


def suite_ec_s():
    return unittest.TestLoader().loadTestsFromTestCase(EmailCollector_select)


if __name__ == "__main__":
    unittest.TextTestRunner(verbosity=2).run(suite_ec())
    unittest.TextTestRunner(verbosity=2).run(suite_ec_s())

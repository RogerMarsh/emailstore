# test_emailcollector.py
# Copyright 2014 Roger Marsh
# Licence: See LICENCE (BSD licence)

"""emailcollector tests"""

import unittest
import os

from .. import emailcollector


class _OperaEmailClient(unittest.TestCase):
    def setUp(self):
        pass

    def tearDown(self):
        pass

    def test___init__01(self):
        oec = emailcollector._OperaEmailClient()
        self.assertEqual(
            oec.mailstore,
            os.path.expanduser(os.path.join("~", ".opera", "mail", "store")),
        )
        self.assertEqual(
            oec.accountdefs,
            os.path.expanduser(
                os.path.join("~", ".opera", "mail", "accounts.ini")
            ),
        )
        self.assertEqual(oec.earliestdate, None)
        self.assertEqual(oec.mostrecentdate, None)
        self.assertEqual(oec.accounts, None)
        self.assertEqual(oec.emailsfrom, None)
        self.assertEqual(oec.outputdirectory, None)
        self.assertEqual(oec._selected_emails, None)
        self.assertEqual(oec._selected_emails_text, None)
        self.assertEqual(oec._filename_map, None)
        self.assertEqual(oec.exclude, None)
        self.assertEqual(len(oec.__dict__), 11)

    def test_get_emails_01(self):
        oec = emailcollector._OperaEmailClient(
            mailstore=("~", "testoperamail", "store"),
            accountdefs=("~", "testoperamail", "accounts.ini"),
        )
        storepath = os.path.expanduser(
            os.path.join("~", "testoperamail", "store")
        )
        self.assertEqual(oec.mailstore, storepath)
        emails = oec.get_emails()
        self.assertIsInstance(emails, list)
        for m in emails:
            self.assertIsInstance(m, tuple)
            self.assertEqual(len(m), 6)
            self.assertEqual(m[0], storepath)
            self.assertEqual(len(m[2]), 4)
            self.assertEqual(len(m[3]), 2)
            self.assertEqual(len(m[4]), 2)
            self.assertTrue(m[2].isdigit())
            self.assertTrue(m[3].isdigit())
            self.assertTrue(m[4].isdigit())
            n, e = os.path.splitext(m[5])
            self.assertEqual(e, ".mbs")
            self.assertTrue(n.isdigit())
            self.assertEqual(m[1][:7], "account")
            self.assertTrue(m[1][7:].isdigit())

    def test_get_emails_02(self):
        oec = emailcollector._OperaEmailClient(
            mailstore=("~", "testoperamail", "store"),
            accountdefs=("~", "testoperamail", "accounts.ini"),
            earliestdate="2012-09-19",
        )
        storepath = os.path.expanduser(
            os.path.join("~", "testoperamail", "store")
        )
        self.assertEqual(oec.mailstore, storepath)
        emails = oec.get_emails()
        self.assertIsInstance(emails, list)
        for m in emails:
            self.assertGreaterEqual("-".join(m[2:5]), "2012-09-19")

    def test_get_emails_03(self):
        oec = emailcollector._OperaEmailClient(
            mailstore=("~", "testoperamail", "store"),
            accountdefs=("~", "testoperamail", "accounts.ini"),
            mostrecentdate="2014-03-25",
        )
        storepath = os.path.expanduser(
            os.path.join("~", "testoperamail", "store")
        )
        self.assertEqual(oec.mailstore, storepath)
        emails = oec.get_emails()
        self.assertIsInstance(emails, list)
        for m in emails:
            self.assertLessEqual("-".join(m[2:5]), "2014-03-25")

    def test_get_emails_04(self):
        oec = emailcollector._OperaEmailClient(
            mailstore=("~", "testoperamail", "store"),
            accountdefs=("~", "testoperamail", "accounts.ini"),
            earliestdate="2009-10-03",
            mostrecentdate="2011-12-20",
        )
        storepath = os.path.expanduser(
            os.path.join("~", "testoperamail", "store")
        )
        self.assertEqual(oec.mailstore, storepath)
        emails = oec.get_emails()
        self.assertIsInstance(emails, list)
        for m in emails:
            self.assertGreaterEqual("-".join(m[2:5]), "2009-10-03")
            self.assertLessEqual("-".join(m[2:5]), "2011-12-20")

    def test_get_emails_05(self):
        oec = emailcollector._OperaEmailClient(
            mailstore=("~", "testoperamail", "store"),
            accountdefs=("~", "testoperamail", "accounts.ini"),
            earliestdate="2013-10-03",
            mostrecentdate="2014-06-20",
            accounts={"roger.marsh@btinternet.com"},
        )
        storepath = os.path.expanduser(
            os.path.join("~", "testoperamail", "store")
        )
        self.assertEqual(oec.mailstore, storepath)
        emails = oec.get_emails()
        self.assertIsInstance(emails, list)
        for m in emails:
            self.assertGreaterEqual("-".join(m[2:5]), "2013-10-03")
            self.assertLessEqual("-".join(m[2:5]), "2014-06-20")

    def test_get_emails_06(self):
        oec = emailcollector._OperaEmailClient(
            mailstore=("~", "testoperamail", "store"),
            accountdefs=("~", "testoperamail", "accounts.ini"),
            earliestdate="2013-10-03",
            mostrecentdate="2014-06-20",
            accounts={"r.m@rmswch.plus.com"},
        )
        storepath = os.path.expanduser(
            os.path.join("~", "testoperamail", "store")
        )
        self.assertEqual(oec.mailstore, storepath)
        emails = oec.get_emails()
        self.assertIsInstance(emails, list)
        for m in emails:
            self.assertGreaterEqual("-".join(m[2:5]), "2013-10-03")
            self.assertLessEqual("-".join(m[2:5]), "2014-06-20")

    def test_get_accounts_01(self):
        oec = emailcollector._OperaEmailClient(
            accountdefs=("~", "testoperamail", "accounts.ini"),
        )
        mailaccountspath = os.path.expanduser(
            os.path.join("~", "testoperamail", "accounts.ini")
        )
        self.assertEqual(oec.accountdefs, mailaccountspath)
        accounts = oec.get_accounts()
        self.assertIsInstance(accounts, dict)
        for k, v in accounts.items():
            self.assertIsInstance(k, str)
            self.assertIsInstance(v, str)
            self.assertEqual(k[:7], "account")
            self.assertTrue(k[7:].isdigit())

    def test_get_emails_for_from_addressees_01(self):
        oec = emailcollector._OperaEmailClient(
            mailstore=("~", "testoperamail", "store"),
            accountdefs=("~", "testoperamail", "accounts.ini"),
            earliestdate="2013-10-03",
            mostrecentdate="2014-06-20",
            accounts={"r.m@rmswch.plus.com"},
        )
        self.assertEqual(
            oec.get_emails(), oec._get_emails_for_from_addressees()
        )

    def test_get_emails_for_from_addressees_02(self):
        oec = emailcollector._OperaEmailClient(
            mailstore=("~", "testoperamail", "store"),
            accountdefs=("~", "testoperamail", "accounts.ini"),
            earliestdate="2013-10-03",
            mostrecentdate="2014-06-20",
            accounts={"r.m@rmswch.plus.com"},
            emailsfrom={},
        )
        self.assertGreaterEqual(
            len(oec.get_emails()), len(oec._get_emails_for_from_addressees())
        )

    def test_get_emails_for_from_addressees_03(self):
        oec = emailcollector._OperaEmailClient(
            mailstore=("~", "testoperamail", "store"),
            accountdefs=("~", "testoperamail", "accounts.ini"),
            earliestdate="2013-10-03",
            mostrecentdate="2014-06-20",
            accounts={"r.m@rmswch.plus.com"},
            emailsfrom={"mcclarkeeastl@yahoo.co.uk"},
        )
        self.assertGreaterEqual(
            len(oec.get_emails()), len(oec._get_emails_for_from_addressees())
        )


class _OperaEmailClient_copy(unittest.TestCase):
    def setUp(self):
        self.opd = os.path.join("~", "testoperacopy")
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

    def test_copy_emails_to_directory_01(self):
        oec = emailcollector._OperaEmailClient(
            mailstore=("~", "testoperamail", "store"),
            accountdefs=("~", "testoperamail", "accounts.ini"),
            earliestdate="2013-11-03",
            mostrecentdate="2014-03-03",
            accounts={"r.m@rmswch.plus.com", "roger.marsh@btinternet.com"},
            emailsfrom={"mcclarkeeastl@yahoo.co.uk"},
            outputdirectory=os.path.expanduser(self.opd),
        )
        exist, equal, copied, changed, exclude = oec.copy_emails_to_directory()
        self.assertEqual(len(equal), 0)
        self.assertEqual(len(changed), 0)
        (
            exist1,
            equal1,
            copied1,
            changed1,
            exclude1,
        ) = oec.copy_emails_to_directory()
        self.assertEqual(len(copied1), 0)
        self.assertEqual(len(changed1), 0)
        self.assertEqual(equal1, copied)
        self.assertEqual(
            len(os.listdir(os.path.expanduser(self.opd))), len(exist1)
        )
        oec = emailcollector._OperaEmailClient(
            mailstore=("~", "testoperamail", "store"),
            accountdefs=("~", "testoperamail", "accounts.ini"),
            earliestdate="2014-02-03",
            mostrecentdate="2014-06-20",
            accounts={"r.m@rmswch.plus.com", "roger.marsh@btinternet.com"},
            emailsfrom={"mcclarkeeastl@yahoo.co.uk"},
            outputdirectory=os.path.expanduser(self.opd),
        )
        (
            exist2,
            equal2,
            copied2,
            changed2,
            exclude2,
        ) = oec.copy_emails_to_directory()
        self.assertEqual(len(changed2), 0)
        self.assertEqual(
            len(os.listdir(os.path.expanduser(self.opd))),
            len(exist2) + len(copied2),
        )
        self.assertEqual(len(copied1) + len(exist1), len(exist2))
        oec = emailcollector._OperaEmailClient(
            mailstore=("~", "testoperamail", "store"),
            accountdefs=("~", "testoperamail", "accounts.ini"),
            earliestdate="2013-10-03",
            mostrecentdate="2013-12-03",
            accounts={"r.m@rmswch.plus.com", "roger.marsh@btinternet.com"},
            emailsfrom={"mcclarkeeastl@yahoo.co.uk"},
            outputdirectory=os.path.expanduser(self.opd),
        )
        (
            exist3,
            equal3,
            copied3,
            changed3,
            exclude3,
        ) = oec.copy_emails_to_directory()
        self.assertEqual(len(changed3), 0)
        self.assertEqual(
            len(os.listdir(os.path.expanduser(self.opd))),
            len(exist3) + len(copied3),
        )
        self.assertEqual(len(copied2) + len(exist2), len(exist3))
        oec = emailcollector._OperaEmailClient(
            mailstore=("~", "testoperamail", "store"),
            accountdefs=("~", "testoperamail", "accounts.ini"),
            earliestdate="2013-11-03",
            mostrecentdate="2014-03-03",
            accounts={"r.m@rmswch.plus.com", "roger.marsh@btinternet.com"},
            emailsfrom={"john.wheeler@care4free.net"},
            outputdirectory=os.path.expanduser(self.opd),
        )
        (
            exist4,
            equal4,
            copied4,
            changed4,
            exclude4,
        ) = oec.copy_emails_to_directory()
        self.assertEqual(len(equal4), 0)
        self.assertEqual(len(changed4), 0)
        self.assertEqual(
            len(os.listdir(os.path.expanduser(self.opd))), len(exist4)
        )
        self.assertEqual(len(copied3) + len(exist3), len(exist4))
        self.assertNotEqual(
            len(copied4), 0, msg="Incomplete test: none to copy."
        )


class EmailCollector(unittest.TestCase):
    def setUp(self):
        pass

    def tearDown(self):
        pass

    def test__assumptions(self):
        msg = "Failure of this test invalidates all other tests"
        self.assertEqual(emailcollector._OPERA_EMAIL_CLIENT, "opera", msg=msg)

    def test___init__01(self):
        ec = emailcollector.EmailCollector()
        self.assertEqual(ec.configuration, None)
        self.assertEqual(ec.dryrun, True)
        self.assertEqual(ec.criteria, None)
        self.assertEqual(ec.email_client, None)
        self.assertEqual(len(ec.__dict__), 4)

    def test_parse_01(self):
        ec = emailcollector.EmailCollector(
            configuration="\n".join(
                (
                    "mailboxstyle opera",
                    "#operamailstore ~/.opera/mail/store",
                    "#operaaccountdefs ~/.opera/mail/accounts.ini",
                    "earliestfromdate 15 April 2011",
                    "mostrecentfromdate 2013-06-25",
                    "account anybody@beeteeinternut.com",
                    "account someoneelse@minus.com",
                    "emailsfrom a.sender@verdant.net",
                    "emailsfrom b.sender@verdant.net",
                    "outputdirectory ~/selected_emails",
                )
            )
        )
        self.assertEqual(ec.parse(), True)

    def test_parse_02(self):
        ec = emailcollector.EmailCollector(
            configuration="\n".join(
                (
                    "       mailboxstyle     opera",
                    "#operamailstore ~/.opera#/mail/store",
                    "    #operaaccountdefs ~/.opera/mail/accounts.ini",
                    "earliestfromdate 15 April 2011#dfdfdfdfd",
                    "earliestfromdate 15    April  2011#dfdfdfdfd",
                    "",
                    "                  ",
                    "account someoneels#### ##   #e@minus.com",
                    "emailsfrom# a.sender@verdant.net",
                    "emailsf#rom b.sender@verdant.net",
                    "outputdirectory #~/selected_emails",
                )
            )
        )
        self.assertEqual(ec.parse(), False)


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
            configuration="\n".join(
                (
                    "mailboxstyle opera",
                    "operamailstore ~/testoperamail/store",
                    "operaaccountdefs ~/testoperamail/accounts.ini",
                    "earliestfromdate 2014-02-03",
                    "mostrecentfromdate 2014-06-20",
                    "account r.m@rmswch.plus.com",
                    "account roger.marsh@btinternet.com",
                    "emailsfrom mcclarkeeastl@yahoo.co.uk",
                    outputdirectory,
                )
            )
        )
        self.assertEqual(ec.parse(), True)
        self.assertEqual(ec.email_client, None)
        self.assertEqual(ec.selected_emails, ec.selected_emails)
        self.assertEqual(ec.selected_emails_text, ec.selected_emails_text)


def suite__oec():
    return unittest.TestLoader().loadTestsFromTestCase(_OperaEmailClient)


def suite__oec_c():
    return unittest.TestLoader().loadTestsFromTestCase(_OperaEmailClient_copy)


def suite_ec():
    return unittest.TestLoader().loadTestsFromTestCase(EmailCollector)


def suite_ec_s():
    return unittest.TestLoader().loadTestsFromTestCase(EmailCollector_select)


if __name__ == "__main__":
    unittest.TextTestRunner(verbosity=2).run(suite__oec())
    unittest.TextTestRunner(verbosity=2).run(suite__oec_c())
    unittest.TextTestRunner(verbosity=2).run(suite_ec())
    unittest.TextTestRunner(verbosity=2).run(suite_ec_s())

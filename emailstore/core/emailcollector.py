# emailcollector.py
# Copyright 2014 Roger Marsh
# Licence: See LICENCE (BSD licence)

"""Retrieve emails from an email client's data store.

Both file, such as mbox, and directory, such as Maildir, formats are used to
store emails.  The file formats put many emails in one file while the directory
formats put one email in each file.

This module supports the mbox file format and the directory format used by the
Opera email client.

It is assumed all email clients are able to export their emails in mbox format.

"""

import os
from datetime import date
import re
from email import message_from_binary_file
from email.utils import parseaddr, parsedate_tz
from email.message import EmailMessage
from mailbox import mbox, mboxMessage, NoSuchMailboxError
import filecmp
from time import strftime
from io import BytesIO
from email.generator import BytesGenerator
import tkinter.messagebox

from solentware_misc.core.utilities import AppSysDate


# The name of the configuration file for selecting emails from a mbox.
COLLECTED_CONF = "collected.conf"

_MBOX_FORMAT = "mbox"
_OPERA_EMAIL_CLIENT = "opera"
_MAILBOX_STYLE = "mailboxstyle"
_OPERA_MAIL_STORE = "operamailstore"
_MBOX_MAIL_STORE = "mboxmailstore"
_OPERA_ACCOUNT_DEFS = "operaaccountdefs"
_EARLIEST_FROM_DATE = "earliestfromdate"
_MOST_RECENT_FROM_DATE = "mostrecentfromdate"
_ACCOUNT = "account"
_EMAILS_FROM = "emailsfrom"
COLLECTED = "collected"
EXCLUDE_EMAIL = "exclude"
_CONF_KEYWORDS = {
    _MAILBOX_STYLE: (_MAILBOX_STYLE, None),
    _OPERA_MAIL_STORE: ("mailstore", None),
    _MBOX_MAIL_STORE: ("mailstore", set),
    _OPERA_ACCOUNT_DEFS: ("accountdefs", None),
    _EARLIEST_FROM_DATE: ("earliestdate", None),
    _MOST_RECENT_FROM_DATE: ("mostrecentdate", None),
    _ACCOUNT: ("accounts", set),
    _EMAILS_FROM: ("emailsfrom", set),
    COLLECTED: ("collected", None),
    EXCLUDE_EMAIL: (EXCLUDE_EMAIL, set),
}


class EmailCollectorError(Exception):
    pass


class EmailCollector(object):

    """Extract emails matching selection criteria from email client store.

    By default look for emails sent or received using the Opera email client
    in the most recent twelve months.

    """

    email_select_line = re.compile(
        "".join(
            (
                "\A",
                "(?:",
                "(?:",  # whitespace line
                "\s*",
                ")|",
                "(?:",  # comment line
                "\s*#.*",
                ")|",
                "(?:",  # parameter line
                "\s*(\S+)\s+([^#]*).*",
                ")",
                ")",
                "\Z",
            )
        )
    )

    def __init__(
        self, directory, configuration=None, dryrun=True, parent=None
    ):
        """Define the email extraction rules from configuration

        directory - the directory containing the configuration file
        configuration - the rules for extracting emails
        dryrun - True: report proposed actions
                 False; do proposed actions after confirmation
        parent - parent widget for dialogues

        """
        self.directory = directory
        self.configuration = configuration
        self.dryrun = dryrun
        self.parent = parent
        self.criteria = None
        self.email_client = None

    def parse(self):
        """ """
        self.criteria = None
        new_values = {k: v() for k, v in _CONF_KEYWORDS.values() if v}
        args = {}
        for line in self.configuration.split("\n"):
            match = self.email_select_line.match(line)
            if not match:
                return False
            key, value = match.groups()
            if key is None:
                continue
            if not value:
                return False
            args_key, args_type = _CONF_KEYWORDS.get(key.lower(), (None, None))
            if args_key is None:
                return False
            elif args_type is None:
                args[args_key] = value
            elif args_type is set:
                args.setdefault(args_key, new_values[args_key]).add(value)
        self.criteria = args
        return True

    def _select_emails(self):
        """ """
        if self.criteria is None:
            return
        if _MAILBOX_STYLE not in self.criteria:
            return
        if self.criteria[_MAILBOX_STYLE].lower() == _OPERA_EMAIL_CLIENT:
            self.email_client = _OperaEmailClient(
                self.directory, self.parent, **self.criteria
            )
        elif self.criteria[_MAILBOX_STYLE].lower() == _MBOX_FORMAT:
            self.email_client = _MboxEmail(
                self.directory, self.parent, **self.criteria
            )
        else:
            return
        return self.email_client.selected_emails

    @property
    def selected_emails(self):
        """ """
        if self.email_client:
            return self.email_client.selected_emails
        return self._select_emails()

    @property
    def selected_emails_text(self):
        """ """
        if self.email_client:
            return self.email_client.selected_emails_text
        elif self._select_emails():
            return self.email_client.selected_emails_text

    @property
    def excluded_emails(self):
        """ """
        if not self.email_client:
            if not self._select_emails():
                return
        return self.email_client.excluded_emails

    @property
    def outputdirectory(self):
        """ """
        return self.email_client.outputdirectory

    @property
    def filename_map(self):
        """ """
        if not self.email_client:
            if not self._select_emails():
                return
        return self.email_client.filename_map

    def copy_emails(self):
        """ """
        if not self.email_client:
            if not self._select_emails():
                return
        return self.email_client.copy_emails_to_directory()

    def exclude_email(self, filename):
        """ """
        if self.email_client.exclude is None:
            self.email_client.exclude = set()
        self.email_client.exclude.add(filename)

    def include_email(self, filename):
        """ """
        if self.email_client.exclude is None:
            self.email_client.exclude = set()
        self.email_client.exclude.remove(filename)


class _MessageFile(EmailMessage):

    """Extend EmailMessage class with a method to generate a filename.

    The From and Date headers are used.

    """

    def generate_filename(self):
        """Return a base filename or None when headers are no available."""
        from_date = parsedate_tz(self.get("Date"))
        from_addr = parseaddr(self.get("From"))[-1]
        if from_date and from_addr:
            date_time_str = strftime("%Y%m%d%H%M%S", from_date[:-1])
            utc = "".join((format(from_date[-1] // 3600, "0=+3"), "00"))
            return "".join((date_time_str, from_addr, utc, ".mbs"))
        else:
            return False


class _MboxMessageFile(mboxMessage):

    """Extend mboxMessage class with a method to generate a filename.

    The From and Date headers are used.

    """

    def generate_filename(self):
        """Return a base filename or None when headers are no available."""
        from_date = parsedate_tz(self.get("Date"))
        from_addr = parseaddr(self.get("From"))[-1]
        if from_date and from_addr:
            date_time_str = strftime("%Y%m%d%H%M%S", from_date[:-1])
            utc = "".join((format(from_date[-1] // 3600, "0=+3"), "00"))
            return "".join((date_time_str, from_addr, utc, ".mbs"))
        else:
            return False


class _OperaEmailClient(object):

    """Extract emails matching selection criteria from Opera email client.

    By default look for emails sent or received in the most recent twelve
    months.

    Opera stores one email per file.  The email is prefixed by a mbox style
    "From " line but lines starting "From " within the email are not converted
    to lines starting ">From ".  So we use the Parser interface rather than the
    Mailbox interface to process the emails.

    """

    # accounts.ini has a number of sections, significant parts here being:
    #
    # ...
    # [account4]
    # ...
    # Email=<email address>
    # ...
    # [account...]
    # ...

    account_ini_line = re.compile(
        b"".join(
            (
                b"\A",
                b"(?:",
                b"(?:",  # [account4]
                b"\[(Account[0-9]+)\]",
                b")|",
                b"(?:",  # [<any other section>]
                b"\[(.*)\]",
                b")|",
                b"(?:",  # Email=<email address>
                b"Email=(.+)",
                b")",
                b")",
                b"\n\Z",
            )
        )
    )

    def __init__(
        self,
        directory,
        parent,
        mailstore=None,
        accountdefs=None,
        accounts=None,
        earliestdate=None,
        mostrecentdate=None,
        emailsfrom=None,
        collected=None,
        exclude=None,
        mailboxstyle=_OPERA_EMAIL_CLIENT,
    ):
        """Define the email extraction rules from configuration

        mailstore - root directory of tree containing email files
        accountdefs - opera file which defines email accounts
        accounts - iterable of email addresses of accounts to be searched
        earliestdate - emails before this date are ignored
        mostrecentdate - emails after this date are ignored
        emailsfrom - iterable of from addressees to select emails
        collected - directory to which email files are copied
        exclude - iterable of email filenames to be ignored when copying
        mailboxstyle - must be 'opera' ignoring case

        See AppSysDate for accepted date formats.  Preferred are '30 Nov 2006'
        and '2006-11-30'.

        """
        self.parent = parent
        if mailboxstyle.lower() != _OPERA_EMAIL_CLIENT:
            raise EmailCollectorError("Mailbox style expected to be Opera")
        if mailstore is None:
            mailstore_path = os.path.join("~", ".opera", "mail", "store")
        elif isinstance(mailstore, (str, bytes)):
            mailstore_path = mailstore
        else:
            mailstore_path = os.path.join(*mailstore)
        if accountdefs is None:
            accountdefs = os.path.join("~", ".opera", "mail", "accounts.ini")
        elif isinstance(accountdefs, (str, bytes)):
            accountdefs = accountdefs
        else:
            accountdefs = os.path.join(*accountdefs)
        self.mailstore = os.path.expanduser(os.path.expandvars(mailstore_path))
        self.accountdefs = os.path.expanduser(os.path.expandvars(accountdefs))
        self.accounts = accounts
        asd = AppSysDate()
        if earliestdate is not None:
            if asd.parse_date(earliestdate) == -1:
                raise EmailCollectorError(
                    "Format error in earliest date argument."
                )
            self.earliestdate = asd.iso_format_date()
        else:
            self.earliestdate = earliestdate
        if mostrecentdate is not None:
            if asd.parse_date(mostrecentdate) == -1:
                raise EmailCollectorError(
                    "Format error in most recent date argument."
                )
            self.mostrecentdate = asd.iso_format_date()
        else:
            self.mostrecentdate = mostrecentdate
        self.emailsfrom = emailsfrom
        if collected == None:
            tkinter.messagebox.showinfo(
                parent=self.parent,
                title="Collect Emails",
                message="".join(
                    (
                        "\n\nDirectory for collected emails not specified:\n\n",
                        "using '",
                        COLLECTED,
                        "' by default.\n\n",
                    )
                ),
            )
            collected = COLLECTED
        self.outputdirectory = os.path.join(directory, collected)
        self.exclude = exclude
        self._selected_emails = None
        self._selected_emails_text = None
        self._filename_map = None

    def get_emails(self):
        """Return email files in order stored in mail store.

        Emails are organized by date and account but each has a filename of
        digits where most recently stored file has highest number.  One email
        per file.

        """
        if self.earliestdate is not None:
            try:
                ymd = tuple([int(d) for d in self.earliestdate.split("-")])
                date(*ymd)
            except:
                raise EmailCollectorError("Earliest date format error")
        else:
            ymd = None
        if self.mostrecentdate is not None:
            try:
                mrd = tuple([int(d) for d in self.mostrecentdate.split("-")])
                date(*mrd)
            except:
                raise EmailCollectorError("Most recent date format error")
        else:
            mrd = None

        # Each email for an account is stored in a file named:
        # <self.mailstore>/<account>/yyyy/mm/dd/<digits>.mbs
        # <account> is a name like 'account3' associated with the email address
        # of the account holder.  An ordered combination of suffix digits not
        # used again if an account is deleted.
        # int(<digits>) > unique integer where n2 stored after n1 if n2 > n1
        # A (send date, sender) is assumed to refer to one file.

        emails = []
        try:
            mailstore = self.mailstore
            accounts = self.get_accounts()
            for account in os.listdir(mailstore):

                # Ignore directories not mentioned in accounts.ini
                if account not in accounts:
                    continue

                # Ignore directories for email accounts not in self.accounts
                if self.accounts:
                    if accounts[account] not in self.accounts:
                        continue

                # Recalculate most recent date for each account if the
                # mostrecentdate argument in _OperaEmailClient() was None
                amrd = mrd
                aed = ymd

                years = sorted(os.listdir(os.path.join(mailstore, account)), reverse=True)
                for year in years:
                    for month in sorted(
                        os.listdir(os.path.join(mailstore, account, year)), reverse=True
                    ):
                        for day in sorted(
                            os.listdir(os.path.join(mailstore, account, year, month)), reverse=True
                        ):
                            if amrd is None:
                                amrd = tuple([int(v) for v in (year, month, day)])
                            if aed is None:
                                aed = tuple([int(year) - 1, int(month), int(day)])
                            emd = (int(year), int(month), int(day))
                            if emd < aed:
                                break
                            if emd > amrd:
                                continue
                            emails.extend(
                                [
                                    (len(e), e, (mailstore, account, year, month, day, e))
                                    for e in os.listdir(
                                        os.path.join(mailstore, account, year, month, day)
                                    )
                                ]
                            )
                        else:
                            continue
                        break
                    else:
                        continue
                    break
        except EmailCollectorError:
            raise
        except:
            if len(emails):
                raise EmailCollectorError(
                    "".join(
                        (
                            "Exception after collecting email ",
                            os.path.join(*emails[-1]),
                        )
                    )
                )
            else:
                raise EmailCollectorError(
                    "Exception before any emails collected."
                )
        emails.sort()
        return [e[-1] for e in emails]

    def get_accounts(self):
        """Return account names associated with owner's email addresses."""
        account_map = {}
        looking_for_email = False
        open_file = open(self.accountdefs, "rb")
        try:
            for line in open_file:
                match = self.account_ini_line.match(line)
                if match:
                    account, header, email = match.groups()
                    if account:
                        if looking_for_email:
                            raise EmailCollectorError(
                                "Unable to map email addresses to accounts"
                            )
                        looking_for_email = True

                        # The account header name is title case but directory
                        # name is lower case (Account5 and account5).
                        account_name = account.decode("utf8").lower()

                    if header and looking_for_email:
                        raise EmailCollectorError(
                            "Unable to map email addresses to accounts"
                        )
                    if email:
                        if not looking_for_email:
                            raise EmailCollectorError(
                                "Unable to map email addresses to accounts"
                            )
                        looking_for_email = False
                        account_map[account_name] = email.decode("utf8")
        finally:
            open_file.close()
        return account_map

    def _get_emails_for_from_addressees(self):
        """Return selected email files in order stored in mail store.

        Emails are selected by 'From Adressee' using the email addresses in
        the emailsfrom argument of _OperaEmailClient() call.

        """
        if self.emailsfrom is None:
            return self.get_emails()
        accounts = self.get_accounts()
        emails = []
        filenamemap = {}
        for email in self.get_emails():
            filename = self._is_from_addressee_of_email_in_selection(email, accounts)
            if filename:
                emails.append(email)
                filenamemap[email[-1]] = filename
        self._filename_map = filenamemap
        return emails

    def _is_from_addressee_of_email_in_selection(self, emailfile, accounts):
        """ """
        if self.emailsfrom is None:
            return True
        file_open = open(os.path.join(*emailfile), "rb")
        try:
            message = message_from_binary_file(file_open, _class=_MessageFile)
            from_ = parseaddr(message.get("From"))[-1]

            # Ignore emails sent by account owner.
            if from_ == accounts[emailfile[1]]:
                return False

            if not self.emailsfrom:
                return message.generate_filename()

            # Ignore emails not sent by someone in self.emailsfrom.
            # Account owners may be in that set, so emails sent from one
            # account owner to another can get selected.
            if from_ in self.emailsfrom:
                return message.generate_filename()
            else:
                return False

        finally:
            file_open.close()

    def copy_emails_to_directory(self):
        """ """
        copied = set()
        changed = set()
        equal = set()
        exist_and_exclude = set()
        directory = self.outputdirectory
        if not os.path.exists(directory):
            os.makedirs(directory)
        exist = set(os.listdir(directory))
        emailfiles = set(self.selected_emails)
        filenamemap = self._filename_map
        exclude = set() if self.exclude is None else self.exclude
        while emailfiles:
            emailpath = emailfiles.pop()
            filename = filenamemap[emailpath[-1]]
            if filename in exclude:
                if filename in exist:
                    exist_and_exclude.add(emailpath)
                continue
            if filename not in exist:
                copied.add(emailpath)
                continue
            if not filecmp.cmp(
                os.path.join(*emailpath),
                os.path.join(directory, filename),
                shallow=False,
            ):
                changed.add(os.path.join(*emailpath))
                continue
            equal.add(emailpath)

        if exist:

            # Change to any files copied previously is sufficient reason to not
            # do any copying at all.
            if changed:
                tkinter.messagebox.showinfo(
                    parent=self.parent,
                    title="Copy Emails to Output Directory",
                    message="".join(
                        (
                            "No emails copied because at least one existing ",
                            "file is different from the file to be copied.",
                        )
                    ),
                )
                return

            # Existence of any file to be excluded is also sufficient reason.
            if exist_and_exclude:
                tkinter.messagebox.showinfo(
                    parent=self.parent,
                    title="Copy Emails to Output Directory",
                    message="".join(
                        (
                            "No emails copied because at least one existing ",
                            "file is currently in the list of files to be ",
                            "excluded from copying.",
                        )
                    ),
                )
                return

            # Merging the existing and copy files cannot be done if the two
            # ranges overlap in sorted order, even if the sets have no files in
            # common.
            if copied:
                filenames = sorted(exist)
                eflow = filenames[0]
                efhigh = filenames[-1]
                sorted_filenames = sorted([filenamemap[e[-1]] for e in copied])
                clow = sorted_filenames[0]
                chigh = sorted_filenames[-1]
                if clow < efhigh:
                    if chigh > eflow:
                        tkinter.messagebox.showinfo(
                            parent=self.parent,
                            title="Copy Emails to Output Directory",
                            message="".join(
                                (
                                    "No emails copied because the range of ",
                                    "file names already in the output ",
                                    "directory overlaps the range of file ",
                                    "names to be copied, when the names are ",
                                    "sorted.\n\nIt is expected file names ",
                                    "start with a datetime formatted to ",
                                    "sort in age order.",
                                )
                            ),
                        )
                        return

        for emailpath in copied:
            input_open = open(os.path.join(*emailpath), "rb")
            try:
                file_open = open(os.path.join(directory, filenamemap[emailpath[-1]]), "wb")
                try:
                    file_open.write(input_open.read())
                finally:
                    file_open.close()
            except FileNotFoundError as exc:
                excdir = os.path.basename(os.path.dirname(exc.filename))
                tkinter.messagebox.showinfo(
                    parent=self.parent,
                    title="Update Extracted Text",
                    message="".join(
                        (
                            "Write additional file to directory\n\n",
                            os.path.basename(os.path.dirname(exc.filename)),
                            "\n\nfailed.\n\nHopefully because the directory ",
                            "does not exist yet: it could have been deleted.",
                        )
                    ),
                )
            finally:
                input_open.close()
        return len(copied)

    @property
    def selected_emails(self):
        """ """
        if self._selected_emails is None:
            self._selected_emails = self._get_emails_for_from_addressees()
        return self._selected_emails

    @property
    def selected_emails_text(self):
        """ """
        if self._selected_emails_text:
            return self._selected_emails_text
        emails_text = []
        for email_filepath in self._selected_emails:
            file_open = open(os.path.join(*email_filepath), "rb")
            try:
                emails_text.append(
                    message_from_binary_file(file_open, _class=_MessageFile)
                )
            finally:
                file_open.close()
        self._selected_emails_text = emails_text
        return self._selected_emails_text

    @property
    def excluded_emails(self):
        """ """
        if not self.exclude:
            return set()
        return set(self.exclude)

    @property
    def filename_map(self):
        """ """
        if not self._filename_map:
            return dict()
        return self._filename_map


class _MboxEmail(object):

    """Extract emails matching selection criteria from a mbox format file.

    By default look for emails sent or received in the most recent twelve
    months.

    Emails are stored in a file in mailbox format.  Most, probably all, email
    clients provide an export function which produces such files if mailbox
    is not the native format of the client.

    """

    def __init__(
        self,
        directory,
        parent,
        mailstore=None,
        accountdefs=None,
        accounts=None,
        earliestdate=None,
        mostrecentdate=None,
        emailsfrom=None,
        collected=None,
        exclude=None,
        mailboxstyle=_MBOX_FORMAT,
    ):
        """Define the email extraction rules from configuration

        mailstore - set of files containining the emails
        accountdefs - ingnored
        accounts - ignored
        earliestdate - emails before this date are ignored
        mostrecentdate - emails after this date are ignored
        emailsfrom - iterable of from addressees to select emails
        collected - directory to which email files are copied
        exclude - iterable of email filenames to be ignored when copying
        mailboxstyle - must be 'mailbox' ignoring case

        See AppSysDate for accepted date formats.  Preferred are '30 Nov 2006'
        and '2006-11-30'.

        """
        self.parent = parent
        if mailboxstyle.lower() != _MBOX_FORMAT:
            raise EmailCollectorError("Mailbox style expected to be mbox")
        if mailstore is None:
            raise EmailCollectorError(
                "The mbox file set is not specified in mailstore argument"
            )
        self.mailstore = set()
        for email in mailstore:
            if isinstance(email, (str, bytes)):
                self.mailstore.add(os.path.expanduser(os.path.expandvars(email)))
            else:
                self.mailstore.add(
                    os.path.expanduser(os.path.expandvars(os.path.join(*email)))
                )
        appsysdate = AppSysDate()
        if earliestdate is None:
            self.earliestdate = earliestdate
        elif appsysdate.parse_date(earliestdate) == -1:
            self.earliestdate = earliestdate
        else:
            self.earliestdate = appsysdate.iso_format_date()
        if mostrecentdate is None:
            self.mostrecentdate = mostrecentdate
        elif appsysdate.parse_date(mostrecentdate) == -1:
            self.mostrecentdate = mostrecentdate
        else:
            self.mostrecentdate = appsysdate.iso_format_date()
        self.emailsfrom = emailsfrom
        if collected == None:
            tkinter.messagebox.showinfo(
                parent=self.parent,
                title="Collect Emails",
                message="".join(
                    (
                        "\n\nDirectory for collected emails not specified:\n\n",
                        "using '",
                        COLLECTED,
                        "' by default.\n\n",
                    )
                ),
            )
            collected = COLLECTED
        self.outputdirectory = os.path.join(directory, collected)
        self.exclude = exclude
        self._selected_emails = None
        self._selected_emails_text = None
        self._filename_map = None

    def get_emails(self):
        """Return messages in order stored in mail store."""

        if self.earliestdate is not None:
            try:
                earliest_date = [d for d in self.earliestdate.split("-")]
                date(*tuple([int(d) for d in earliest_date]))
                earliest_date = "".join(earliest_date)
            except:
                tkinter.messagebox.showinfo(
                    parent=self.parent,
                    title="Select Emails",
                    message="".join(
                        (
                            "\n\nDate for earliest emails to be selected\n\n",
                            str(self.earliestdate),
                            "\n\nis not in a correct format.",
                        )
                    ),
                )
                return []
        else:
            earliest_date = None
        if self.mostrecentdate is not None:
            try:
                mrd = [d for d in self.mostrecentdate.split("-")]
                date(*tuple([int(d) for d in mrd]))
                mrd = "".join(mrd)
            except:
                tkinter.messagebox.showinfo(
                    parent=self.parent,
                    title="Select Emails",
                    message="".join(
                        (
                            "\n\nDate for most recent emails to be selected",
                            "\n\n",
                            str(self.mostrecentdate),
                            "\n\nis not in a correct format.",
                        )
                    ),
                )
                return []
        else:
            mrd = None

        # All emails are stored in the files named in self.mailstore.
        # Get all the email objects and sort into 'sent date' order.
        # A (send date, sender) is assumed to refer to one email which may
        # be present in more than one mbox-style file.

        emails = {}
        timefrom = {}
        try:
            for mailstore in self.mailstore:
                try:
                    mailstore_mbox = mbox(
                        mailstore, factory=_MboxMessageFile, create=False
                    )
                except NoSuchMailboxError as exc:
                    tkinter.messagebox.showinfo(
                        parent=self.parent,
                        title="Mailbox Not Found",
                        message="".join(
                            (
                                "File\n\n",
                                os.path.basename(str(exc)),
                                "\n\ndoes not exist.\n\nAny emails found ",
                                "in other files have been ignored.",
                            )
                        ),
                    )
                    return []
                for message in mailstore_mbox.itervalues():
                    filename = message.generate_filename()
                    fnd = filename[:8]
                    if earliest_date is not None:
                        if fnd < earliest_date:
                            continue
                    if mrd is not None:
                        if fnd > mrd:
                            continue
                    msgid = message.get("Message-ID")
                    if filename not in timefrom:
                        timefrom[filename] = set()
                    timefrom[filename].add(msgid)

                    # Assume it is impossible two different emails have same
                    # timestamp, from addressee, and message-id.
                    emails[(filename, msgid)] = message

        except EmailCollectorError:
            raise
        except:
            if len(emails):
                raise EmailCollectorError(
                    "".join(
                        (
                            "Exception after collecting email ",
                            os.path.join(*emails[-1]),
                        )
                    )
                )
            else:
                raise EmailCollectorError(
                    "Exception before any emails collected."
                )
        for k, value in timefrom.items():
            if len(value) == 1:
                emails[k] = emails.pop((k, value.pop()))
            else:
                for i in value:
                    emails["".join((k, i))] = emails.pop((k, i))
        emails = [(k, v) for k, v in emails.items()]
        emails.sort()
        return emails

    def _get_emails_for_from_addressees(self):
        """Return selected email files in order stored in mail store.

        Emails are selected by 'From Adressee' using the email addresses in
        the emailsfrom argument of _OperaEmailClient() call.

        """
        if self.emailsfrom is None:
            return self.get_emails()
        emails = []
        for email in self.get_emails():
            fntrue = self._is_from_addressee_of_email_in_selection(email[-1])
            if fntrue:
                emails.append(email)
        return emails

    def _is_from_addressee_of_email_in_selection(self, emailfile):
        """ """
        if self.emailsfrom is None:
            return True

        # By analogy with _OperaEmailClient version of this method
        from_ = parseaddr(emailfile.get("From"))[-1]
        if not self.emailsfrom:
            return True

        # Ignore emails not sent by someone in self.emailsfrom.
        return bool(from_ in self.emailsfrom)

    def copy_emails_to_directory(self):
        """ """
        copied = set()
        changed = set()
        equal = set()
        exist_and_exclude = set()
        directory = self.outputdirectory
        if not os.path.exists(directory):
            os.makedirs(directory)
        exist = set(os.listdir(directory))
        emailfiles = set(self.selected_emails)
        exclude = set() if self.exclude is None else self.exclude
        while emailfiles:
            filename, message = emailfiles.pop()
            if filename in exclude:
                if filename in exist:
                    exist_and_exclude.add(message)
                continue
            if filename not in exist:
                copied.add((filename, message))
                continue

            # message is in memory as mboxMessage so read (directory, filename)
            bytes_io = BytesIO()
            generator = BytesGenerator(bytes_io, mangle_from_=False, maxheaderlen=0)
            generator.flatten(message)
            text = bytes_io.getvalue()
            if (
                bytes_io.getvalue()
                != open(os.path.join(directory, filename), "rb").read()
            ):
                changed.add(filename)
                continue
            equal.add(message)

        if exist:

            # Change to any files copied previously is sufficient reason to not
            # do any copying at all.
            if changed:
                tkinter.messagebox.showinfo(
                    parent=self.parent,
                    title="Copy Emails to Output Directory",
                    message="".join(
                        (
                            "No emails copied because at least one existing ",
                            "file is different from the file to be copied.",
                        )
                    ),
                )
                return

            # Existence of any file to be excluded is also sufficient reason.
            if exist_and_exclude:
                tkinter.messagebox.showinfo(
                    parent=self.parent,
                    title="Copy Emails to Output Directory",
                    message="".join(
                        (
                            "No emails copied because at least one existing ",
                            "file is currently in the list of files to be ",
                            "excluded from copying.",
                        )
                    ),
                )
                return

            # Merging the existing and copy files cannot be done if the two
            # ranges overlap in sorted order, even if the sets have no files in
            # common.
            if copied:
                filenames = sorted(exist)
                eflow = filenames[0]
                efhigh = filenames[-1]
                sorted_filenames = sorted([e[0] for e in copied])
                clow = sorted_filenames[0]
                chigh = sorted_filenames[-1]
                if clow < efhigh:
                    if chigh > eflow:
                        tkinter.messagebox.showinfo(
                            parent=self.parent,
                            title="Copy Emails to Output Directory",
                            message="".join(
                                (
                                    "No emails copied because the range of ",
                                    "file names already in the output ",
                                    "directory overlaps the range of file ",
                                    "names to be copied, when the names are ",
                                    "sorted.\n\nIt is expected file names ",
                                    "start with a datetime formatted to ",
                                    "sort in age order.",
                                )
                            ),
                        )
                        return

        for filename, message in copied:
            bytes_io = BytesIO()
            generator = BytesGenerator(bytes_io, mangle_from_=False, maxheaderlen=0)
            generator.flatten(message)
            text = bytes_io.getvalue()
            try:
                file_open = open(os.path.join(directory, filename), "wb")
            except FileNotFoundError as exc:
                excdir = os.path.basename(os.path.dirname(exc.filename))
                tkinter.messagebox.showinfo(
                    parent=self.parent,
                    title="Copy Emails to Output Directory",
                    message="".join(
                        (
                            "Write additional file to directory\n\n",
                            os.path.basename(os.path.dirname(exc.filename)),
                            "\n\nfailed.\n\nHopefully because the directory ",
                            "does not exist yet: it could have been deleted.",
                        )
                    ),
                )
                return
            try:
                file_open.write(text)
            finally:
                file_open.close()
        return len(copied)

    @property
    def selected_emails(self):
        """ """
        if self._selected_emails is None:
            self._selected_emails = self._get_emails_for_from_addressees()
        return self._selected_emails

    @property
    def selected_emails_text(self):
        """ """
        if self._selected_emails_text:
            return self._selected_emails_text
        emails_text = []
        for filename, message in self._selected_emails:
            emails_text.append(message)
        self._selected_emails_text = emails_text
        return self._selected_emails_text

    @property
    def excluded_emails(self):
        """ """
        if not self.exclude:
            return set()
        return set(self.exclude)

    @property
    def filename_map(self):
        """ """
        if not self._filename_map:
            return dict()
        return self._filename_map

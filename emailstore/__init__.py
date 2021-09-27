# __init__.py
# Copyright 2014 Roger Marsh
# Licence: See LICENCE (BSD licence)

"""Select emails from an email client's mailboxes and store in a directory.

Email clients use a variety of formats for mailboxes.  Some use a file as a
mailbox: these hold many emails per file.  Others use a directory as a mailbox:
these hold one email per file.  Some directory formats allow both files and
directories to exist in a directory while others limit the content to either
files or directories.

Emailstore flattens the mailbox naming hierarchy and stores *.mbs files in a
directory, relying on the email client's naming conventions to avoid conflict.

The *.mbs files are pruned using the selection criteria.  When a select and
store action is repeated the previously stored version must exist exactly in
the version about to replace it.  'old_email_str in new_email_str' must be True
expresses the idea in Python code.

Currently the Opera version of mailbox directories is the only style supported.

"""
APPLICATION_NAME = "EmailStore"
ERROR_LOG = "ErrorLog"

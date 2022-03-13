# select.py
# Copyright 2014 Roger Marsh
# Licence: See LICENCE (BSD licence)

"""Email selection filter User Interface.

"""

import os
import tkinter
import tkinter.messagebox
import tkinter.filedialog
from email.utils import parseaddr, parsedate_tz
from time import strftime

from solentware_misc.gui.exceptionhandler import ExceptionHandler
from solentware_misc.gui import textreadonly
from solentware_misc.gui.configuredialog import ConfigureDialog

from . import help
from .. import APPLICATION_NAME
from ..core.emailcollector import (
    EmailCollector,
    EXCLUDE_EMAIL,
    COLLECTED_CONF,
    _MBOX_MAIL_STORE,
    _MAILBOX_STYLE,
)

startup_minimum_width = 340
startup_minimum_height = 400


class SelectError(Exception):
    pass


class Select(ExceptionHandler):

    """Define and use an email select and store configuration file."""

    def __init__(
        self,
        folder=None,
        use_toplevel=False,
        application_name=APPLICATION_NAME,
        **kargs
    ):
        """Create the database and GUI objects.

        **kargs - passed to tkinter Toplevel widget if use_toplevel True

        """
        if use_toplevel:
            self.root = tkinter.Toplevel(**kargs)
        else:
            self.root = tkinter.Tk()
        try:
            if folder is not None:
                self.root.wm_title(" - ".join((application_name, folder)))
            else:
                self.root.wm_title(application_name)
            self.root.wm_minsize(
                width=startup_minimum_width, height=startup_minimum_height
            )
            self.application_name = application_name

            self._configuration = None
            self._configuration_edited = False
            self._email_collector = None
            self._tag_names = set()
            self._excluded = set()

            menubar = tkinter.Menu(self.root)

            menufile = tkinter.Menu(menubar, name="file", tearoff=False)
            menubar.add_cascade(label="File", menu=menufile, underline=0)
            menufile.add_command(
                label="Open",
                underline=0,
                command=self.try_command(self.file_open, menufile),
            )
            menufile.add_command(
                label="New",
                underline=0,
                command=self.try_command(self.file_new, menufile),
            )
            menufile.add_separator()
            # menufile.add_command(
            #    label='Save',
            #    underline=0,
            #    command=self.try_command(self.file_save, menufile))
            menufile.add_command(
                label="Save Copy As...",
                underline=7,
                command=self.try_command(self.file_save_copy_as, menufile),
            )
            menufile.add_separator()
            menufile.add_command(
                label="Close",
                underline=0,
                command=self.try_command(self.file_close, menufile),
            )
            menufile.add_separator()
            menufile.add_command(
                label="Quit",
                underline=0,
                command=self.try_command(self.file_quit, menufile),
            )

            menuactions = tkinter.Menu(menubar, name="actions", tearoff=False)
            menubar.add_cascade(label="Actions", menu=menuactions, underline=0)
            menuactions.add_command(
                label="Show selection",
                underline=0,
                command=self.try_command(self.show_selection, menuactions),
            )
            menuactions.add_command(
                label="Apply selection",
                underline=0,
                command=self.try_command(self.apply_selection, menuactions),
            )
            menuactions.add_command(
                label="Clear selection",
                underline=0,
                command=self.try_command(self.clear_selection, menuactions),
            )
            menufile.add_separator()
            menuactions.add_command(
                label="Option editor",
                underline=0,
                command=self.try_command(
                    self.configure_email_selection, menuactions
                ),
            )

            menuhelp = tkinter.Menu(menubar, name="help", tearoff=False)
            menubar.add_cascade(label="Help", menu=menuhelp, underline=0)
            menuhelp.add_command(
                label="Guide",
                underline=0,
                command=self.try_command(self.help_guide, menuhelp),
            )
            menuhelp.add_command(
                label="Notes",
                underline=0,
                command=self.try_command(self.help_notes, menuhelp),
            )
            menuhelp.add_command(
                label="About",
                underline=0,
                command=self.try_command(self.help_about, menuhelp),
            )

            self.root.configure(menu=menubar)

            self.statusbar = Statusbar(self.root)
            frame = tkinter.PanedWindow(
                self.root,
                background="cyan2",
                opaqueresize=tkinter.FALSE,
                orient=tkinter.HORIZONTAL,
            )
            frame.pack(fill=tkinter.BOTH, expand=tkinter.TRUE)

            toppane = tkinter.PanedWindow(
                master=frame,
                opaqueresize=tkinter.FALSE,
                orient=tkinter.HORIZONTAL,
            )
            originalpane = tkinter.PanedWindow(
                master=toppane,
                opaqueresize=tkinter.FALSE,
                orient=tkinter.VERTICAL,
            )
            emailpane = tkinter.PanedWindow(
                master=toppane,
                opaqueresize=tkinter.FALSE,
                orient=tkinter.VERTICAL,
            )
            self.configctrl = textreadonly.make_text_readonly(
                master=originalpane, width=80
            )
            self.emaillistctrl = textreadonly.make_text_readonly(
                master=originalpane, width=80
            )
            self.emailtextctrl = textreadonly.make_text_readonly(
                master=emailpane
            )
            originalpane.add(self.configctrl)
            originalpane.add(self.emaillistctrl)
            emailpane.add(self.emailtextctrl)
            toppane.add(originalpane)
            toppane.add(emailpane)
            toppane.pack(side=tkinter.TOP, expand=True, fill=tkinter.BOTH)
            for widget, sequence, function in (
                (self.configctrl, "<ButtonPress-3>", self.conf_popup),
                (self.emaillistctrl, "<ButtonPress-3>", self.list_popup),
                (self.emailtextctrl, "<ButtonPress-3>", self.text_popup),
            ):
                widget.bind(sequence, self.try_event(function))
            mbox_popup_menu = tkinter.Menu(
                master=self.configctrl, tearoff=False
            )
            mbox_popup_menu.add_separator()
            mbox_popup_menu.add_command(
                label="Replace and #comment current",
                command=self.replace_and_comment_current,
                accelerator="Alt F5",
            )
            mbox_popup_menu.add_command(
                label="Replace",
                command=self.replace_current,
                accelerator="Alt F3",
            )
            mbox_popup_menu.add_command(
                label="Insert after current",
                command=self.insert_after_current,
                accelerator="Alt F7",
            )
            mbox_popup_menu.add_command(
                label="Insert before current",
                command=self.insert_before_current,
                accelerator="Alt F9",
            )
            mbox_popup_menu.add_separator()
            self.mbox_popup_menu = mbox_popup_menu
            mboxstyle_popup_menu = tkinter.Menu(
                master=self.configctrl, tearoff=False
            )
            mboxstyle_popup_menu.add_separator()
            mboxstyle_popup_menu.add_command(
                label="Insert after mailboxstyle",
                command=self.insert_after_current,
                accelerator="Alt F7",
            )
            mboxstyle_popup_menu.add_separator()
            self.mboxstyle_popup_menu = mboxstyle_popup_menu
            self._folder = folder
            self._most_recent_action = None

        except:
            self.root.destroy()
            del self.root

    def __del__(self):
        """ """
        if self._configuration:
            self._configuration = None

    def help_about(self):
        """Display information about EmailStore."""
        help.help_about(self.root)

    def help_guide(self):
        """Display brief User Guide for EmailStore."""
        help.help_guide(self.root)

    def help_notes(self):
        """Display technical notes about EmailStore."""
        help.help_notes(self.root)

    def get_toplevel(self):
        """Return the toplevel widget."""
        return self.root

    def file_new(self):
        """Create and open a new email selection configuration file."""
        if self._configuration is not None:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title=self.application_name,
                message="Close the current email selection first.",
            )
            return
        config_file = tkinter.filedialog.asksaveasfilename(
            parent=self.get_toplevel(),
            title=" ".join(("New", self.application_name)),
            defaultextension=".conf",
            filetypes=(("Email Selection Rules", "*.conf"),),
            initialfile=COLLECTED_CONF,
            initialdir=self._folder if self._folder else "~",
        )
        if not config_file:
            return
        self.configctrl.delete("1.0", tkinter.END)
        self.configctrl.insert(
            tkinter.END,
            "".join(
                ("# ", os.path.basename(config_file), " email selection rules")
            ),
        )
        self._save_configuration(set_edited_flag=False)
        self._configuration = config_file
        self._folder = os.path.dirname(config_file)
        self.root.wm_title(" - ".join((self.application_name, config_file)))

    def file_open(self):
        """Open an existing email selection rules file."""
        if self._configuration is not None:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title=self.application_name,
                message="Close the current email selection first.",
            )
            return
        config_file = tkinter.filedialog.askopenfilename(
            parent=self.get_toplevel(),
            title=" ".join(("Open", self.application_name)),
            defaultextension=".conf",
            filetypes=(("Email Selection Rules", "*.conf"),),
            initialfile=COLLECTED_CONF,
            initialdir=self._folder if self._folder else "~",
        )
        if not config_file:
            return
        fn = open(config_file, "r")
        try:
            self.configctrl.delete("1.0", tkinter.END)
            self.configctrl.insert(tkinter.END, fn.read())
        finally:
            fn.close()
        self._configuration = config_file
        self._folder = os.path.dirname(config_file)
        self.root.wm_title(" - ".join((self.application_name, config_file)))

    def file_close(self):
        """Close the open email selection rules file."""
        if self._configuration is None:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title=self.application_name,
                message="Cannot close.\n\nThere is no file open.",
            )
            return
        dlg = tkinter.messagebox.askquestion(
            parent=self.get_toplevel(),
            title=self.application_name,
            message="Confirm Close.",
        )
        if dlg == tkinter.messagebox.YES:
            self._clear_email_tags()
            self.configctrl.delete("1.0", tkinter.END)
            self.emailtextctrl.delete("1.0", tkinter.END)
            self.emaillistctrl.delete("1.0", tkinter.END)
            self.statusbar.set_status_text()
            self._configuration = None
            self._configuration_edited = False
            self._email_collector = None
            self.root.wm_title(
                " - ".join((self.application_name, self._folder))
            )

    def file_quit(self):
        """Quit the email selection application."""
        dlg = tkinter.messagebox.askquestion(
            parent=self.get_toplevel(),
            title=self.application_name,
            message="Confirm Quit.",
        )
        if dlg == tkinter.messagebox.YES:
            self.root.destroy()

    def file_save_copy_as(self):
        """Save copy of open email selection rules and keep current open."""
        if self._configuration is None:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title="Save Copy As",
                message="Cannot save.\n\nEmail selection rules file not open.",
            )
            return
        config_file = tkinter.filedialog.asksaveasfilename(
            parent=self.get_toplevel(),
            title=self.application_name.join(("Save ", " As")),
            defaultextension=".conf",
            filetypes=(("Email Selection Rules", "*.conf"),),
            initialfile=os.path.basename(self._configuration),
            initialdir=os.path.dirname(self._configuration),
        )
        if not config_file:
            return
        if config_file == self._configuration:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title="Save Copy As",
                message="".join(
                    (
                        'Cannot use "Save Copy As" to overwite the open ',
                        "email selection rules file.",
                    )
                ),
            )
            return
        self._save_configuration(set_edited_flag=False)

    def configure_email_selection(self):
        """Set parameters that control email selection from mailboxes."""
        if self._configuration is None:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title="Configure Email Selection",
                message="Open an email selection rules file.",
            )
            return
        config_text = ConfigureDialog(
            self.root,
            configuration=self.configctrl.get(
                "1.0", " ".join((tkinter.END, "-1 chars"))
            ),
            dialog_title=" ".join(
                (self.application_name, "configuration editor")
            ),
        ).config_text
        if config_text is None:
            return
        self._configuration_edited = True
        self.configctrl.delete("1.0", tkinter.END)
        self.configctrl.insert(tkinter.END, config_text)
        fn = open(self._configuration, "w")
        try:
            fn.write(config_text)
            self._clear_email_tags()
            self.emailtextctrl.delete("1.0", tkinter.END)
            self.emaillistctrl.delete("1.0", tkinter.END)
            self.statusbar.set_status_text()
            self._configuration_edited = False
            self._email_collector = None
        finally:
            fn.close()
        if self._most_recent_action:
            self._most_recent_action()

    def show_selection(self):
        """Do the email selection but do not copy the emails."""
        if self._configuration is None:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title="Show Email Selection",
                message="Open an email selection rules file",
            )
            return
        if self._configuration_edited:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title="Show Email Selection",
                message="".join(
                    (
                        "The edited configuration file has not been saved. ",
                        'It must be saved before "Show" action can be done.',
                    )
                ),
            )
            return
        if self._email_collector is None:
            emc = EmailCollector(
                os.path.dirname(self._configuration),
                configuration=self.configctrl.get(
                    "1.0", " ".join((tkinter.END, "-1 chars"))
                ),
                dryrun=True,
                parent=self.get_toplevel(),
            )
            if not emc.parse():
                tkinter.messagebox.showinfo(
                    parent=self.get_toplevel(),
                    title="Show Email Selection",
                    message="Email selection rules are invalid",
                )
                return
            if not emc.selected_emails:
                tkinter.messagebox.showinfo(
                    parent=self.get_toplevel(),
                    title="Show Email Selection",
                    message="No emails match the selection rules.",
                )
                return
            self._email_collector = emc
        self._show_selection()
        self._most_recent_action = self.show_selection
        return True

    def _show_selection(self):
        """Do the email selection but do not copy the emails."""
        self._clear_email_tags()
        tw = self.emailtextctrl
        lw = self.emaillistctrl
        cc = self.configctrl
        tags = self._tag_names
        tw.delete("1.0", tkinter.END)
        lw.delete("1.0", tkinter.END)

        # Tag the text put in the widgets such that the source entry in
        # selected_emails_text can be recovered from the pointer position
        # over the widget.
        try:
            for e, m in enumerate(self._email_collector.selected_emails_text):
                textname = "x".join(("T", str(e)))
                tags.add(textname)
                entryname = "x".join(("M", str(e)))
                tags.add(entryname)
                fromname = "x".join(("F", str(e)))
                tags.add(fromname)
                start = tw.index(tkinter.INSERT)
                tw.insert(tkinter.END, m.as_string())
                tw.insert(tkinter.END, "\n")
                tw.tag_add(textname, start, tw.index(tkinter.INSERT))
                tw.insert(tkinter.END, "\n\n\n")
                tw.tag_add(entryname, start, tw.index(tkinter.INSERT))
                start = lw.index(tkinter.INSERT)
                fromstart = lw.index(tkinter.INSERT)
                lw.insert(tkinter.END, m.get("From", ""))
                lw.insert(tkinter.END, "\n")
                lw.insert(tkinter.END, m.get("Date", ""))
                lw.tag_add(fromname, fromstart, lw.index(tkinter.INSERT))
                lw.insert(tkinter.END, "\n")
                lw.insert(tkinter.END, m.get("Subject", ""))
                lw.insert(tkinter.END, "\n")
                lw.tag_add(textname, start, lw.index(tkinter.INSERT))
                lw.insert(tkinter.END, "\n\n")
                lw.tag_add(entryname, start, lw.index(tkinter.INSERT))
                fm, dt = lw.get(*lw.tag_ranges(fromname)).split("\n")
                dt = parsedate_tz(dt)
                fm = parseaddr(fm)
                date = strftime("%Y%m%d%H%M%S", dt[:-1])
                utc = "".join((format(dt[-1] // 3600, "0=+3"), "00"))
                filename = "".join((date, fm[-1], utc, ".mbs"))
                if filename in self._email_collector.excluded_emails:
                    si = cc.search(filename, "1.0")
                    start = cc.index(" ".join((si, "linestart")))
                    end = cc.index(" ".join((si, "lineend")))
                    cc.tag_add(fromname, start, end)
                    cc.tag_bind(fromname, "<ButtonPress-1>", self._file_exists)
        except TypeError:
            if not self._email_collector.selected_emails_text:
                tkinter.messagebox.showinfo(
                    parent=self.get_toplevel(),
                    title="Email Selection",
                    message="".join(
                        (
                            "No emails found.\n\n",
                            "(Email addresses spelt correctly?)",
                        )
                    ),
                )
                return
            raise

    def apply_selection(self):
        """Do the email selection and copy the emails on confirmation."""
        if not self.show_selection():
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title="Apply Email Selection",
                message="Unable to apply selection",
            )
            return
        if (
            tkinter.messagebox.askquestion(
                parent=self.get_toplevel(),
                title="Apply Email Selection",
                message="".join(
                    (
                        "Confirm request to apply email selection ",
                        "and copy selected emails.",
                    )
                ),
            )
            != tkinter.messagebox.YES
        ):
            return
        ce = self._email_collector.copy_emails()
        if ce is not None:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title="Apply Email Selection",
                message="".join(
                    (
                        "Apply selection done: ",
                        str(ce),
                        " file copied." if ce == 1 else " files copied.",
                    )
                ),
            )
        return

    def clear_selection(self):
        """Clear the lists of selected emails."""
        if (
            tkinter.messagebox.askquestion(
                parent=self.get_toplevel(),
                title="Clear Email Lists",
                message="Confirm request to clear lists of selected emails.",
            )
            != tkinter.messagebox.YES
        ):
            return
        self._clear_email_tags()
        self.emailtextctrl.delete("1.0", tkinter.END)
        self.emaillistctrl.delete("1.0", tkinter.END)
        self.statusbar.set_status_text()
        self._email_collector = None
        self._most_recent_action = None

    def _clear_email_tags(self):
        """Clear the tags identifying data for each email."""
        for w in (self.emailtextctrl, self.emaillistctrl, self.configctrl):
            for t in self._tag_names:
                w.tag_delete(t)
        self._tag_names.clear()

    def conf_popup(self, event=None):
        """ """
        wconf = self.configctrl
        index = wconf.index("".join(("@", str(event.x), ",", str(event.y))))
        start = wconf.index(" ".join((index, "linestart")))
        end = wconf.index(" ".join((index, "lineend", "+1 char")))
        text = wconf.get(start, end)
        keyvalue = text.split(" ", maxsplit=1)
        if len(keyvalue) == 1:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title="Edit Email Selection",
                message="Popup menu not supported for 'key only' lines",
            )
            return
        if keyvalue[0] not in (
            EXCLUDE_EMAIL,
            _MBOX_MAIL_STORE,
            _MAILBOX_STYLE,
        ):
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title="Edit Email Selection",
                message="".join(
                    (
                        "Popup menu not supported for '",
                        keyvalue[0],
                        "' lines",
                    )
                ),
            )
            return
        if keyvalue[0] == EXCLUDE_EMAIL:
            if (
                tkinter.messagebox.askquestion(
                    parent=self.get_toplevel(),
                    title="Cancel Exclude Email",
                    message="".join(
                        (
                            "Confirm request to cancel exclusion of \n\n",
                            text.split(" ", 1)[-1],
                            "\n\nemail.\n\nThe file is not copied to the ",
                            "output directory in this action; use ",
                            '"Apply" later to do this.',
                        )
                    ),
                )
                != tkinter.messagebox.YES
            ):
                return
            wconf.delete(start, end)
            if self._email_collector is not None:
                self._email_collector.include_email(
                    text.split(" ", 1)[-1].strip()
                )
            self._save_configuration()
            return
        if keyvalue[0] == _MBOX_MAIL_STORE:
            mboxpath = os.path.expanduser(keyvalue[1])
            while mboxpath not in ("/", ""):
                if os.path.isdir(mboxpath):
                    break
                mboxpath = os.path.dirname(mboxpath)
            if mboxpath in ("/", ""):
                tkinter.messagebox.showinfo(
                    parent=self.get_toplevel(),
                    title="Add or Amend 'mailstore'",
                    message="".join(
                        (
                            "Directory '",
                            os.path.dirname(keyvalue[1]),
                            "' not found. Defaulting to user's  ",
                            "home directory.",
                        )
                    ),
                )
                mboxpath = os.path.expanduser("~")
            self.__mboxpath = mboxpath
            self.__start = start
            self.__end = end
            self.mbox_popup_menu.tk_popup(*event.widget.winfo_pointerxy())
            return
        if keyvalue[0] == _MAILBOX_STYLE:
            self.__mboxpath = os.path.expanduser("~")
            self.__end = end
            self.mboxstyle_popup_menu.tk_popup(*event.widget.winfo_pointerxy())
            return
        tkinter.messagebox.showinfo(
            parent=self.get_toplevel(),
            title=" ".join((keyvalue[0], " not Supported")),
            message="".join(
                (
                    "Add or amend '",
                    keyvalue[0],
                    "' not yet implemented.  Edit manually",
                )
            ),
        )
        return

    def replace_and_comment_current(self, event=None):
        """ """
        mbox_filepath = self.get_new_mbox_filepath(
            "Replace and Comment Current"
        )
        if not mbox_filepath:
            return
        self.configctrl.insert(self.__start, "#")
        self.configctrl.insert(
            self.__end + "-1 char",
            "".join(("\n", mbox_filepath)),
        )
        self._save_configuration()
        return

    def replace_current(self, event=None):
        """ """
        mbox_filepath = self.get_new_mbox_filepath("Replace Current")
        if not mbox_filepath:
            return
        self.configctrl.delete(self.__start, self.__end + "-1 char")
        self.configctrl.insert(self.__start, mbox_filepath)
        self._save_configuration()
        return

    def insert_after_current(self, event=None):
        """ """
        mbox_filepath = self.get_new_mbox_filepath("Insert after Current")
        if not mbox_filepath:
            return
        self.configctrl.insert(
            self.__end + "-1 char",
            "".join(("\n", mbox_filepath)),
        )
        self._save_configuration()
        return

    def insert_before_current(self, event=None):
        """ """
        mbox_filepath = self.get_new_mbox_filepath("Insert before Current")
        if not mbox_filepath:
            return
        self.configctrl.insert(
            self.__start,
            "".join((mbox_filepath, "\n")),
        )
        self._save_configuration()
        return

    def get_new_mbox_filepath(self, title=""):
        """ """
        filepath = tkinter.filedialog.askopenfilename(
            parent=self.get_toplevel(),
            title=title,
            initialdir=self.__mboxpath,
        )
        del self.__mboxpath
        if not filepath:
            tkinter.messagebox.showinfo(
                parent=self.get_toplevel(),
                title=title,
                message="Action not done: no filename selected",
            )
        home = os.path.expanduser("~")
        if filepath.startswith("".join((home, "/"))):
            filepath = filepath.replace(home, "~")
        return " ".join((_MBOX_MAIL_STORE, filepath))

    def list_popup(self, event=None):
        """ """
        wtext = self.emailtextctrl
        wlist = self.emaillistctrl
        wconf = self.configctrl
        tags = wlist.tag_names(
            wlist.index("".join(("@", str(event.x), ",", str(event.y))))
        )
        for t in tags:
            if t.startswith("F"):
                text = wlist.get(*wlist.tag_ranges(t))
                if (
                    tkinter.messagebox.askquestion(
                        parent=self.get_toplevel(),
                        title="Show Email in List",
                        message="".join(
                            (
                                "Confirm request to scroll text to \n\n",
                                text,
                                "\n\nemail.",
                            )
                        ),
                    )
                    != tkinter.messagebox.YES
                ):
                    return
                wtext.see(wtext.tag_ranges("".join(("T", t[1:])))[0])
                trconf = wconf.tag_ranges(t)
                if trconf:
                    wconf.see(trconf[-1])
                return

    def text_popup(self, event=None):
        """ """
        wtext = self.emailtextctrl
        wlist = self.emaillistctrl
        tags = wtext.tag_names(
            wtext.index("".join(("@", str(event.x), ",", str(event.y))))
        )
        for t in tags:
            if t.startswith("T"):
                ftag = "".join(("F", t[1:]))
                fm, dt = wlist.get(*wlist.tag_ranges(ftag)).split("\n")
                dt = parsedate_tz(dt)
                fm = parseaddr(fm)
                if not (fm and dt):
                    tkinter.messagebox.showinfo(
                        parent=self.get_toplevel(),
                        title="Remove Email from Selection",
                        message="Email from or date invalid.",
                    )
                    return
                date = strftime("%Y%m%d%H%M%S", dt[:-1])
                utc = "".join((format(dt[-1] // 3600, "0=+3"), "00"))
                filename = "".join((date, fm[-1], utc, ".mbs"))
                if filename in self._email_collector.excluded_emails:
                    tkinter.messagebox.showinfo(
                        parent=self.get_toplevel(),
                        title="Remove Email from Selection",
                        message="".join(
                            (
                                filename,
                                "\n\n",
                                "is already one of the emails excluded from ",
                                "the selection.",
                            )
                        ),
                    )
                    return
                fp = os.path.join(
                    os.path.expanduser(self._email_collector.outputdirectory),
                    filename,
                )
                if os.path.exists(fp):
                    if (
                        tkinter.messagebox.askquestion(
                            parent=self.get_toplevel(),
                            title="Remove Email from Selection",
                            message="".join(
                                (
                                    filename,
                                    "\n\nexists in the output directory.  ",
                                    "You will have to use your system's file ",
                                    "manager to delete the file.\n\nConfirm ",
                                    "request to add \n\n",
                                    wlist.get(*wlist.tag_ranges(ftag)),
                                    "\n\nto exclude email list in selection ",
                                    "rules.",
                                )
                            ),
                        )
                        != tkinter.messagebox.YES
                    ):
                        return
                elif (
                    tkinter.messagebox.askquestion(
                        parent=self.get_toplevel(),
                        title="Remove Email from Selection",
                        message="".join(
                            (
                                "Confirm request to add \n\n",
                                wlist.get(*wlist.tag_ranges(ftag)),
                                "\n\nto exclude email list in selection ",
                                "rules.",
                            )
                        ),
                    )
                    != tkinter.messagebox.YES
                ):
                    return
                wconf = self.configctrl
                start = wconf.index(tkinter.END)
                wconf.insert(tkinter.END, "\n")
                wconf.insert(tkinter.END, " ".join((EXCLUDE_EMAIL, filename)))
                wconf.tag_add(
                    ftag, start, wconf.index(" ".join((start, "lineend")))
                )
                wconf.tag_bind(ftag, "<ButtonPress-1>", self._file_exists)
                self._email_collector.exclude_email(filename)
                self._save_configuration()
                return

    def _save_configuration(self, set_edited_flag=True):
        """ """
        if set_edited_flag:
            self._configuration_edited = True
        fn = open(self._configuration, "w")
        try:
            fn.write(
                self.configctrl.get("1.0", " ".join((tkinter.END, "-1 chars")))
            )
            if set_edited_flag:
                self._configuration_edited = False
        finally:
            fn.close()
        self._clear_email_tags()
        self.emailtextctrl.delete("1.0", tkinter.END)
        self.emaillistctrl.delete("1.0", tkinter.END)
        self.statusbar.set_status_text()
        self._email_collector = None
        if self._most_recent_action:
            self._most_recent_action()

    def _file_exists(self, event=None):
        """ """
        w = event.widget
        ti = w.index("".join(("@", str(event.x), ",", str(event.y))))
        start = w.index(" ".join((ti, "linestart")))
        end = w.index(" ".join((ti, "lineend")))
        filename = w.get(start, end).split(" ", 1)[-1]
        od = os.path.expanduser(self._email_collector.outputdirectory)
        fp = os.path.join(od, filename)
        if os.path.exists(fp):
            self.statusbar.set_status_text(
                " ".join((filename, "exists in output directory", od))
            )
        else:
            self.statusbar.set_status_text(
                " ".join((filename, "does not exist in output directory", od))
            )


class Statusbar(object):

    """Status bar for EmailStore application."""

    def __init__(self, root):
        """Create status bar widget."""
        self.status = tkinter.Text(
            root,
            height=0,
            width=0,
            background=root.cget("background"),
            relief=tkinter.FLAT,
            state=tkinter.DISABLED,
            wrap=tkinter.NONE,
        )
        self.status.pack(side=tkinter.BOTTOM, fill=tkinter.X)

    def get_status_text(self):
        """Return text displayed in status bar."""
        return self.status.cget("text")

    def set_status_text(self, text=""):
        """Display text in status bar."""
        self.status.configure(state=tkinter.NORMAL)
        self.status.delete("1.0", tkinter.END)
        self.status.insert(tkinter.END, text)
        self.status.configure(state=tkinter.DISABLED)
